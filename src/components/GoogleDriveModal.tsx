import React, { useEffect, useState } from 'react';
import {
  AlertTriangle,
  Check,
  Cloud,
  CloudDownload,
  CloudUpload,
  Copy,
  ExternalLink,
  FileText,
  FolderOpen,
  FolderSync,
  Globe,
  Link2,
  LogOut,
  RefreshCw,
  Search,
  Share2,
  Trash2,
  X,
} from 'lucide-react';
import { User } from 'firebase/auth';
import {
  deleteFileFromDrive,
  downloadDriveDocxAsArrayBuffer,
  downloadDriveJsonFile,
  DriveFileItem,
  extractDriveFileIdFromInput,
  getAccessToken,
  listDriveWordAndBackupFiles,
  publishLibraryToSharedGoogleDrive,
  SHARED_BACKUP_FILE_NAME,
  SHARED_DRIVE_FOLDER_NAME,
} from '../utils/googleDriveService';
import {
  fetchStaticBundledBackup,
  loadPublicDriveLibraryByFileId,
  SharedDrivePublicationInfo,
  TemplatesConfigMap,
} from '../utils/templateStore';

export const GoogleSignInButton: React.FC<{
  onClick: () => void;
  disabled?: boolean;
  label?: string;
}> = ({ onClick, disabled = false, label = 'Sign in with Google' }) => {
  return (
    <button
      type="button"
      onClick={onClick}
      disabled={disabled}
      className="gsi-material-button shrink-0"
    >
      <div className="gsi-material-button-state"></div>
      <div className="gsi-material-button-content-wrapper">
        <div className="gsi-material-button-icon">
          <svg
            version="1.1"
            xmlns="http://www.w3.org/2000/svg"
            viewBox="0 0 48 48"
            xmlnsXlink="http://www.w3.org/1999/xlink"
            style={{ display: 'block' }}
          >
            <path
              fill="#EA4335"
              d="M24 9.5c3.54 0 6.71 1.22 9.21 3.6l6.85-6.85C35.9 2.38 30.47 0 24 0 14.62 0 6.51 5.38 2.56 13.22l7.98 6.19C12.43 13.72 17.74 9.5 24 9.5z"
            ></path>
            <path
              fill="#4285F4"
              d="M46.98 24.55c0-1.57-.15-3.09-.38-4.55H24v9.02h12.94c-.58 2.96-2.26 5.48-4.78 7.18l7.73 6c4.51-4.18 7.09-10.36 7.09-17.65z"
            ></path>
            <path
              fill="#FBBC05"
              d="M10.53 28.59c-.48-1.45-.76-2.99-.76-4.59s.27-3.14.76-4.59l-7.98-6.19C.92 16.46 0 20.12 0 24c0 3.88.92 7.54 2.56 10.78l7.97-6.19z"
            ></path>
            <path
              fill="#34A853"
              d="M24 48c6.48 0 11.93-2.13 15.89-5.81l-7.73-6c-2.15 1.45-4.92 2.3-8.16 2.3-6.26 0-11.57-4.22-13.47-9.91l-7.98 6.19C6.51 42.62 14.62 48 24 48z"
            ></path>
            <path fill="none" d="M0 0h48v48H0z"></path>
          </svg>
        </div>
        <span className="gsi-material-button-contents">{label}</span>
        <span style={{ display: 'none' }}>{label}</span>
      </div>
    </button>
  );
};

interface GoogleDriveModalProps {
  isOpen: boolean;
  onClose: () => void;
  user: User | null;
  needsAuth: boolean;
  isLoggingIn: boolean;
  onLogin: () => Promise<void>;
  onLogout: () => Promise<void>;
  templatesConfig: TemplatesConfigMap;
  sharedDriveInfo: SharedDrivePublicationInfo | null;
  onUpdateSharedDriveInfo: (
    info: SharedDrivePublicationInfo,
    updatedTemplates: TemplatesConfigMap
  ) => void;
  onImportDocxFromDrive: (file: File) => Promise<void>;
  onRestoreLibraryFromDrive: (restoredConfig: TemplatesConfigMap) => void;
  onAuthExpired: () => void;
}

interface PendingConfirmAction {
  type: 'delete_file' | 'publish_shared_drive';
  file?: DriveFileItem;
  title: string;
  description: string;
}

export const GoogleDriveModal: React.FC<GoogleDriveModalProps> = ({
  isOpen,
  onClose,
  user,
  needsAuth,
  isLoggingIn,
  onLogin,
  onLogout,
  templatesConfig,
  sharedDriveInfo,
  onUpdateSharedDriveInfo,
  onImportDocxFromDrive,
  onRestoreLibraryFromDrive,
  onAuthExpired,
}) => {
  const [driveFiles, setDriveFiles] = useState<DriveFileItem[]>([]);
  const [searchQuery, setSearchQuery] = useState<string>('');
  const [isLoadingFiles, setIsLoadingFiles] = useState<boolean>(false);
  const [busyFileId, setBusyFileId] = useState<string | null>(null);
  const [isPublishingToDrive, setIsPublishingToDrive] = useState<boolean>(false);
  const [publicDriveLinkInput, setPublicDriveLinkInput] = useState<string>('');
  const [isLinkingPublicDrive, setIsLinkingPublicDrive] = useState<boolean>(false);
  const [copiedShareLink, setCopiedShareLink] = useState<boolean>(false);
  const [statusMessage, setStatusMessage] = useState<{
    type: 'success' | 'error' | 'info';
    text: string;
  } | null>(null);

  // Mandatory confirmation dialog state for destructive/mutating Drive actions
  const [pendingConfirm, setPendingConfirm] =
    useState<PendingConfirmAction | null>(null);

  const fetchFiles = async (query = searchQuery) => {
    if (needsAuth || !user) return;
    setIsLoadingFiles(true);
    try {
      const files = await listDriveWordAndBackupFiles(query);
      setDriveFiles(files);
    } catch (err: any) {
      if (err?.message === 'UNAUTHENTICATED') {
        onAuthExpired();
      } else {
        setStatusMessage({
          type: 'error',
          text:
            err instanceof Error
              ? err.message
              : 'Không thể tải danh sách tệp từ Google Drive.',
        });
      }
    } finally {
      setIsLoadingFiles(false);
    }
  };

  useEffect(() => {
    if (isOpen && !needsAuth && user) {
      fetchFiles('');
    }
  }, [isOpen, needsAuth, user]);

  if (!isOpen) return null;

  const existingBackupFile = driveFiles.find(
    (f) => f.name === SHARED_BACKUP_FILE_NAME
  );

  // Handle importing a .docx or Google Doc from Drive into the Template Configurator
  const handleSelectDocxFromDrive = async (item: DriveFileItem) => {
    setBusyFileId(item.id);
    setStatusMessage(null);
    try {
      const { buffer, fileName } = await downloadDriveDocxAsArrayBuffer(item);
      const docxFile = new File([buffer], fileName, {
        type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      });
      await onImportDocxFromDrive(docxFile);
      onClose();
    } catch (err: any) {
      if (err?.message === 'UNAUTHENTICATED') {
        onAuthExpired();
      } else {
        setStatusMessage({
          type: 'error',
          text:
            err instanceof Error
              ? err.message
              : 'Lỗi khi tải biểu mẫu từ Google Drive.',
        });
      }
    } finally {
      setBusyFileId(null);
    }
  };

  // Handle restoring templates_store_backup.json from Drive
  const handleRestoreFromBackupFile = async (item: DriveFileItem) => {
    setBusyFileId(item.id);
    setStatusMessage(null);
    try {
      const jsonText = await downloadDriveJsonFile(item.id);
      const parsed = JSON.parse(jsonText);
      if (!parsed || typeof parsed !== 'object') {
        throw new Error('Nội dung tệp sao lưu không hợp lệ.');
      }
      onRestoreLibraryFromDrive(parsed as TemplatesConfigMap);
      setStatusMessage({
        type: 'success',
        text: `Đã đồng bộ và khôi phục thư viện mẫu từ tệp "${item.name}" trên Google Drive!`,
      });
    } catch (err: any) {
      if (err?.message === 'UNAUTHENTICATED') {
        onAuthExpired();
      } else {
        setStatusMessage({
          type: 'error',
          text:
            err instanceof Error
              ? err.message
              : 'Không thể khôi phục từ tệp sao lưu trên Drive.',
        });
      }
    } finally {
      setBusyFileId(null);
    }
  };

  // Publish & Share the entire template library on Google Drive so ANY app user can access it
  const executePublishAndShareOnDrive = async () => {
    setIsPublishingToDrive(true);
    setStatusMessage(null);
    try {
      const targetBackupId =
        sharedDriveInfo?.backupFileId || existingBackupFile?.id;

      const { driveInfo, updatedTemplatesConfig } =
        await publishLibraryToSharedGoogleDrive({
          templatesConfig,
          ownerEmail: user?.email || undefined,
          existingBackupFileId: targetBackupId,
          uploadIndividualDocxFiles: true,
        });

      onUpdateSharedDriveInfo(driveInfo, updatedTemplatesConfig);

      setStatusMessage({
        type: 'success',
        text: `Đã lưu toàn bộ ${driveInfo.templateCount} biểu mẫu lên thư mục "${SHARED_DRIVE_FOLDER_NAME}" trên Google Drive và bật chia sẻ công khai! Bất kỳ ai truy cập ứng dụng đều tự động nhận được dữ liệu từ Google Drive của bạn.`,
      });

      await fetchFiles(searchQuery);
    } catch (err: any) {
      if (err?.message === 'UNAUTHENTICATED') {
        onAuthExpired();
      } else {
        setStatusMessage({
          type: 'error',
          text:
            err instanceof Error
              ? err.message
              : 'Không thể lưu và chia sẻ kho mẫu lên Google Drive.',
        });
      }
    } finally {
      setIsPublishingToDrive(false);
    }
  };

  // Link a public Google Drive Share URL or File ID (works for any user)
  const handleLinkPublicDriveUrl = async () => {
    const fileId = extractDriveFileIdFromInput(publicDriveLinkInput);
    if (!fileId) {
      setStatusMessage({
        type: 'error',
        text: 'Vui lòng dán đường link chia sẻ tệp Google Drive hợp lệ (VD: https://drive.google.com/file/d/.../view?usp=sharing) hoặc File ID.',
      });
      return;
    }

    setIsLinkingPublicDrive(true);
    setStatusMessage(null);
    try {
      const token = await getAccessToken();
      const { driveConfig, templates } = await loadPublicDriveLibraryByFileId(
        fileId,
        token
      );
      onUpdateSharedDriveInfo(driveConfig, templates);
      setPublicDriveLinkInput('');
      setStatusMessage({
        type: 'success',
        text: `Đã kết nối và nạp thành công ${Object.keys(templates).length} biểu mẫu từ Google Drive (File ID: ${fileId}) cho toàn bộ người dùng ứng dụng!`,
      });
    } catch (err) {
      setStatusMessage({
        type: 'error',
        text:
          err instanceof Error
            ? err.message
            : 'Không thể đọc tệp từ đường link Google Drive này.',
      });
    } finally {
      setIsLinkingPublicDrive(false);
    }
  };

  // Import templates_store_backup.json directly from computer (works 100% offline/static on GitHub Pages)
  const handleLocalBackupJsonUpload = async (
    e: React.ChangeEvent<HTMLInputElement>
  ) => {
    const file = e.target.files?.[0];
    e.target.value = '';
    if (!file) return;

    setStatusMessage(null);
    try {
      const text = await file.text();
      const parsed = JSON.parse(text);
      if (!parsed || typeof parsed !== 'object' || Array.isArray(parsed)) {
        throw new Error('Tệp JSON không đúng định dạng kho biểu mẫu.');
      }
      const templates = parsed as TemplatesConfigMap;
      const count = Object.keys(templates).length;
      if (count === 0) {
        throw new Error('Tệp JSON không chứa biểu mẫu nào.');
      }
      const firstWithDrive = Object.values(templates).find((t) => t.drive_file_id);
      const info: SharedDrivePublicationInfo = {
        backupFileId: firstWithDrive?.drive_file_id || 'local_backup_json',
        backupFileName: file.name,
        backupWebViewLink: firstWithDrive?.drive_web_view_link,
        folderName: SHARED_DRIVE_FOLDER_NAME,
        updatedAt: new Date().toISOString(),
        templateCount: count,
      };
      onUpdateSharedDriveInfo(info, templates);
      setStatusMessage({
        type: 'success',
        text: `Đã nạp thành công ${count} biểu mẫu từ tệp "${file.name}"!`,
      });
    } catch (err) {
      setStatusMessage({
        type: 'error',
        text:
          err instanceof Error
            ? err.message
            : 'Không thể đọc tệp JSON sao lưu.',
      });
    }
  };

  // Load the pre-bundled static templates_store_backup.json from GitHub Pages / public folder
  const handleLoadStaticBundledLibrary = async () => {
    setIsLinkingPublicDrive(true);
    setStatusMessage(null);
    try {
      const bundled = await fetchStaticBundledBackup();
      const count = Object.keys(bundled).length;
      if (count === 0) {
        throw new Error('Không tìm thấy tệp kho mẫu tích hợp sẵn.');
      }
      const firstWithDrive = Object.values(bundled).find((t) => t.drive_file_id);
      const info: SharedDrivePublicationInfo = {
        backupFileId: firstWithDrive?.drive_file_id || 'github_pages_bundled',
        backupFileName: SHARED_BACKUP_FILE_NAME,
        backupWebViewLink: firstWithDrive?.drive_web_view_link,
        folderName: SHARED_DRIVE_FOLDER_NAME,
        updatedAt: new Date().toISOString(),
        templateCount: count,
      };
      onUpdateSharedDriveInfo(info, bundled);
      setStatusMessage({
        type: 'success',
        text: `Đã đồng bộ thành công ${count} biểu mẫu từ Kho Mẫu Tích Hợp Sẵn!`,
      });
    } catch (err) {
      setStatusMessage({
        type: 'error',
        text:
          err instanceof Error
            ? err.message
            : 'Không thể nạp kho mẫu tích hợp sẵn.',
      });
    } finally {
      setIsLinkingPublicDrive(false);
    }
  };

  // Execute confirmed destructive/mutating operation
  const handleConfirmPendingAction = async () => {
    if (!pendingConfirm) return;
    const action = pendingConfirm;
    setPendingConfirm(null);

    if (action.type === 'publish_shared_drive') {
      await executePublishAndShareOnDrive();
      return;
    }

    if (action.type === 'delete_file' && action.file) {
      setBusyFileId(action.file.id);
      setStatusMessage(null);
      try {
        await deleteFileFromDrive(action.file.id);
        setDriveFiles((prev) => prev.filter((f) => f.id !== action.file!.id));
        setStatusMessage({
          type: 'success',
          text: `Đã xóa tệp "${action.file.name}" khỏi Google Drive.`,
        });
      } catch (err: any) {
        if (err?.message === 'UNAUTHENTICATED') {
          onAuthExpired();
        } else {
          setStatusMessage({
            type: 'error',
            text:
              err instanceof Error
                ? err.message
                : 'Không thể xóa tệp trên Google Drive.',
          });
        }
      } finally {
        setBusyFileId(null);
      }
    }
  };

  const appShareUrlWithDrive = sharedDriveInfo?.backupFileId
    ? `${window.location.origin}${window.location.pathname}?driveId=${encodeURIComponent(
        sharedDriveInfo.backupFileId
      )}`
    : `${window.location.origin}${window.location.pathname}`;

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/50 px-4 py-6">
      <div className="w-full max-w-3xl bg-white border border-slate-200 rounded-lg overflow-hidden flex flex-col max-h-[90vh]">
        {/* Modal Header */}
        <div className="flex items-center justify-between gap-4 px-6 py-4 bg-slate-50 border-b border-slate-200">
          <div className="flex items-center gap-2.5">
            <Cloud className="w-5 h-5 text-blue-700" />
            <div>
              <h2 className="text-base font-bold text-slate-900 font-display">
                Lưu Trữ & Chia Sẻ Kho Biểu Mẫu Trên Google Drive
              </h2>
              <p className="text-xs text-slate-500">
                Lưu kho mẫu trên Google Drive của bạn và tự động chia sẻ dữ liệu với bất kỳ ai truy cập ứng dụng
              </p>
            </div>
          </div>

          <button
            type="button"
            onClick={onClose}
            className="p-1.5 text-slate-400 hover:text-slate-700 rounded-lg cursor-pointer"
          >
            <X className="w-4 h-4" />
          </button>
        </div>

        {/* Modal Body */}
        <div className="p-6 overflow-y-auto space-y-5 flex-1">
          {/* Active Shared Google Drive Storage Status Card (Visible to ALL users) */}
          <div className="p-4 bg-emerald-50/70 border border-emerald-200 rounded-lg space-y-3">
            <div className="flex flex-wrap items-start justify-between gap-3">
              <div className="space-y-1">
                <div className="flex items-center gap-2">
                  <Globe className="w-4 h-4 text-emerald-700 shrink-0" />
                  <h3 className="text-xs font-bold text-emerald-950 uppercase tracking-wider">
                    Kho Dữ Liệu Google Drive Dùng Chung Cho Mọi Người Truy Cập App
                  </h3>
                </div>
                {sharedDriveInfo ? (
                  <p className="text-xs text-emerald-900 leading-relaxed">
                    Đang kết nối trực tiếp với Google Drive{' '}
                    {sharedDriveInfo.ownerEmail ? (
                      <span className="font-semibold">({sharedDriveInfo.ownerEmail})</span>
                    ) : null}{' '}
                    · Thư mục:{' '}
                    <span className="font-semibold">
                      {sharedDriveInfo.folderName || SHARED_DRIVE_FOLDER_NAME}
                    </span>{' '}
                    · <span className="font-mono-tabular font-semibold">{Object.keys(templatesConfig).length} mẫu</span>
                  </p>
                ) : (
                  <p className="text-xs text-emerald-900 leading-relaxed">
                    Hãy bấm <strong>&ldquo;Lưu & Chia Sẻ Kho Mẫu Lên Google Drive&rdquo;</strong> bên dưới để lưu toàn bộ các biểu mẫu vào Google Drive của bạn và cấp quyền đọc tự động cho tất cả người truy cập ứng dụng.
                  </p>
                )}
              </div>

              {sharedDriveInfo && (
                <div className="flex flex-wrap items-center gap-2 shrink-0">
                  {sharedDriveInfo.folderWebViewLink && (
                    <a
                      href={sharedDriveInfo.folderWebViewLink}
                      target="_blank"
                      rel="noopener noreferrer"
                      className="inline-flex items-center gap-1 px-2.5 py-1.5 text-xs font-semibold text-emerald-900 bg-white border border-emerald-300 hover:bg-emerald-100 rounded-md whitespace-nowrap"
                    >
                      <FolderOpen className="w-3.5 h-3.5 text-emerald-700" />
                      <span>Mở thư mục Drive</span>
                      <ExternalLink className="w-3 h-3" />
                    </a>
                  )}
                  {sharedDriveInfo.backupWebViewLink && (
                    <a
                      href={sharedDriveInfo.backupWebViewLink}
                      target="_blank"
                      rel="noopener noreferrer"
                      className="inline-flex items-center gap-1 px-2.5 py-1.5 text-xs font-medium text-emerald-900 bg-white border border-emerald-300 hover:bg-emerald-100 rounded-md whitespace-nowrap"
                    >
                      <FileText className="w-3.5 h-3.5 text-emerald-700" />
                      <span>File dữ liệu Drive</span>
                      <ExternalLink className="w-3 h-3" />
                    </a>
                  )}
                </div>
              )}
            </div>

            {sharedDriveInfo && (
              <div className="pt-2 border-t border-emerald-200/80 flex flex-col sm:flex-row sm:items-center justify-between gap-2">
                <div className="text-[11px] font-mono-tabular text-emerald-800 truncate">
                  Link chia sẻ App (đọc từ Google Drive): {appShareUrlWithDrive}
                </div>
                <button
                  type="button"
                  onClick={() => {
                    navigator.clipboard.writeText(appShareUrlWithDrive);
                    setCopiedShareLink(true);
                    setTimeout(() => setCopiedShareLink(false), 2500);
                  }}
                  className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold text-white bg-emerald-700 hover:bg-emerald-800 rounded-md whitespace-nowrap shrink-0 cursor-pointer"
                >
                  {copiedShareLink ? (
                    <>
                      <Check className="w-3.5 h-3.5" />
                      <span>Đã copy Link App + Drive!</span>
                    </>
                  ) : (
                    <>
                      <Copy className="w-3.5 h-3.5" />
                      <span>Copy Link Chia Sẻ App</span>
                    </>
                  )}
                </button>
              </div>
            )}
          </div>

          {/* Mandatory Confirmation Dialog for Overwrite / Publishing / Deleting on Google Drive */}
          {pendingConfirm && (
            <div className="p-4 bg-amber-50 border border-amber-300 rounded-lg space-y-3">
              <div className="flex items-start gap-2.5">
                <AlertTriangle className="w-5 h-5 text-amber-700 shrink-0 mt-0.5" />
                <div className="space-y-1">
                  <h4 className="text-xs font-bold text-amber-950">
                    {pendingConfirm.title}
                  </h4>
                  <p className="text-xs text-amber-900 leading-relaxed">
                    {pendingConfirm.description}
                  </p>
                </div>
              </div>
              <div className="flex items-center justify-end gap-2">
                <button
                  type="button"
                  onClick={() => setPendingConfirm(null)}
                  className="px-3 py-1.5 text-xs font-medium text-slate-700 bg-white border border-slate-300 hover:bg-slate-100 rounded-md cursor-pointer"
                >
                  Hủy bỏ
                </button>
                <button
                  type="button"
                  onClick={handleConfirmPendingAction}
                  className={`px-3.5 py-1.5 text-xs font-semibold text-white rounded-md cursor-pointer ${
                    pendingConfirm.type === 'delete_file'
                      ? 'bg-red-600 hover:bg-red-700'
                      : 'bg-blue-700 hover:bg-blue-800'
                  }`}
                >
                  {pendingConfirm.type === 'delete_file'
                    ? 'Xác nhận xóa vĩnh viễn'
                    : 'Xác nhận lưu & chia sẻ trên Google Drive'}
                </button>
              </div>
            </div>
          )}

          {statusMessage && (
            <div
              className={`p-3 rounded-lg text-xs font-medium border ${
                statusMessage.type === 'success'
                  ? 'bg-emerald-50 text-emerald-900 border-emerald-200'
                  : statusMessage.type === 'error'
                  ? 'bg-red-50 text-red-900 border-red-200'
                  : 'bg-blue-50 text-blue-900 border-blue-200'
              }`}
            >
              {statusMessage.text}
            </div>
          )}

          {needsAuth || !user ? (
            <div className="p-6 bg-slate-50 border border-slate-200 rounded-lg text-center space-y-4">
              <Cloud className="w-9 h-9 text-blue-700 mx-auto" />
              <div className="space-y-1 max-w-md mx-auto">
                <h3 className="text-sm font-semibold text-slate-900">
                  Đăng nhập Google để lưu và chia sẻ kho mẫu từ Google Drive của bạn
                </h3>
                <p className="text-xs text-slate-600 leading-relaxed">
                  Khi bạn đăng nhập và bấm &ldquo;Lưu & Chia Sẻ Kho Mẫu Lên Google Drive&rdquo;, toàn bộ các mẫu Word (.docx) và cấu hình trường Mailings sẽ được lưu vào Google Drive của bạn và chia sẻ quyền đọc để bất kỳ ai mở ứng dụng đều truy cập được.
                </p>
              </div>
              <div className="flex justify-center pt-1">
                <GoogleSignInButton
                  onClick={onLogin}
                  disabled={isLoggingIn}
                  label={isLoggingIn ? 'Connecting...' : 'Sign in with Google'}
                />
              </div>
            </div>
          ) : (
            <>
              {/* Connected User Bar & Primary "Save & Share on Google Drive" Action */}
              <div className="p-4 bg-slate-50 border border-slate-200 rounded-lg flex flex-col sm:flex-row sm:items-center justify-between gap-4">
                <div className="flex items-center gap-3 min-w-0">
                  <div className="w-9 h-9 rounded-full bg-blue-700 text-white font-semibold text-sm flex items-center justify-center shrink-0">
                    {(user.displayName || user.email || 'G')[0].toUpperCase()}
                  </div>
                  <div className="min-w-0">
                    <p className="text-xs font-semibold text-slate-900 truncate">
                      {user.displayName || 'Chủ sở hữu Kho Google Drive'}
                    </p>
                    <p className="text-[11px] text-slate-500 truncate">
                      {user.email}
                    </p>
                  </div>
                </div>

                <div className="flex flex-wrap items-center gap-2">
                  <button
                    type="button"
                    disabled={isPublishingToDrive}
                    onClick={() =>
                      setPendingConfirm({
                        type: 'publish_shared_drive',
                        title:
                          'Lưu & Chia sẻ công khai Kho Biểu Mẫu trên Google Drive?',
                        description: `Ứng dụng sẽ tải toàn bộ ${
                          Object.keys(templatesConfig).length
                        } biểu mẫu Word (.docx) cùng tệp cấu hình "${SHARED_BACKUP_FILE_NAME}" vào thư mục "${SHARED_DRIVE_FOLDER_NAME}" trên Google Drive (${
                          user.email
                        }) và bật quyền xem ("Bất kỳ ai có đường liên kết đều có thể xem") để mọi người truy cập ứng dụng đều sử dụng được các mẫu này.`,
                      })
                    }
                    className="inline-flex items-center gap-1.5 px-3.5 py-2 text-xs font-semibold text-white bg-blue-700 hover:bg-blue-800 disabled:opacity-60 rounded-md whitespace-nowrap shrink-0 cursor-pointer"
                  >
                    <Share2 className="w-3.5 h-3.5" />
                    <span>
                      {isPublishingToDrive
                        ? 'Đang lưu & bật chia sẻ trên Google Drive...'
                        : 'Lưu & Chia Sẻ Kho Mẫu Lên Google Drive'}
                    </span>
                  </button>

                  <button
                    type="button"
                    onClick={onLogout}
                    className="inline-flex items-center gap-1 px-2.5 py-2 text-xs font-medium text-slate-600 bg-white border border-slate-300 hover:bg-slate-100 rounded-md whitespace-nowrap shrink-0 cursor-pointer"
                  >
                    <LogOut className="w-3.5 h-3.5" />
                    <span>Đăng xuất</span>
                  </button>
                </div>
              </div>

              {/* Search & Refresh Drive Files */}
              <div className="flex items-center gap-2">
                <div className="relative flex-1">
                  <Search className="w-4 h-4 text-slate-400 absolute left-3 top-1/2 -translate-y-1/2" />
                  <input
                    type="text"
                    value={searchQuery}
                    onChange={(e) => setSearchQuery(e.target.value)}
                    onKeyDown={(e) => {
                      if (e.key === 'Enter') fetchFiles(searchQuery);
                    }}
                    placeholder="Tìm tệp .docx, Google Docs hoặc bản sao lưu trên Google Drive..."
                    className="w-full pl-9 pr-3 py-2 text-xs bg-slate-50 border border-slate-300 rounded-lg text-slate-900 focus:bg-white focus:outline-none focus:border-blue-600"
                  />
                </div>
                <button
                  type="button"
                  onClick={() => fetchFiles(searchQuery)}
                  disabled={isLoadingFiles}
                  className="inline-flex items-center gap-1.5 px-3.5 py-2 text-xs font-semibold text-slate-700 bg-slate-100 hover:bg-slate-200 rounded-lg whitespace-nowrap shrink-0 cursor-pointer"
                >
                  <RefreshCw
                    className={`w-3.5 h-3.5 ${
                      isLoadingFiles ? 'animate-spin' : ''
                    }`}
                  />
                  <span>Tìm / Làm mới</span>
                </button>
              </div>

              {/* Files List */}
              <div className="border border-slate-200 rounded-lg overflow-hidden">
                <div className="px-4 py-2.5 bg-slate-50 border-b border-slate-200 flex items-center justify-between text-xs font-semibold text-slate-600">
                  <span>Tệp Word (.docx), Google Docs & Kho sao lưu trên Drive của bạn</span>
                  <span className="font-mono-tabular">{driveFiles.length} tệp</span>
                </div>

                {isLoadingFiles ? (
                  <div className="p-8 text-center text-xs text-slate-500">
                    Đang tải danh sách tệp từ Google Drive...
                  </div>
                ) : driveFiles.length === 0 ? (
                  <div className="p-8 text-center text-xs text-slate-500">
                    Không tìm thấy tệp Word (.docx) hoặc Google Docs nào khớp với tìm kiếm trên Google Drive.
                  </div>
                ) : (
                  <div className="divide-y divide-slate-200 max-h-[260px] overflow-y-auto">
                    {driveFiles.map((file) => {
                      const isBackupJson =
                        file.name.includes('templates_store_backup') ||
                        file.mimeType.includes('json');
                      const isGoogleDoc =
                        file.mimeType === 'application/vnd.google-apps.document';
                      const isBusy = busyFileId === file.id;

                      return (
                        <div
                          key={file.id}
                          className="px-4 py-3 flex flex-col sm:flex-row sm:items-center justify-between gap-3 hover:bg-slate-50 transition-colors"
                        >
                          <div className="flex items-start gap-2.5 min-w-0">
                            <FileText className="w-4 h-4 text-blue-700 shrink-0 mt-0.5" />
                            <div className="min-w-0">
                              <p className="text-xs font-semibold text-slate-900 truncate">
                                {file.name}
                              </p>
                              <div className="flex flex-wrap items-center gap-1.5 text-[11px] text-slate-500">
                                <span>
                                  {isBackupJson
                                    ? 'Tệp dữ liệu Kho biểu mẫu dùng chung (JSON)'
                                    : isGoogleDoc
                                    ? 'Google Docs (Tự động chuyển sang .docx)'
                                    : 'Microsoft Word (.docx)'}
                                </span>
                                {file.modifiedTime && (
                                  <>
                                    <span aria-hidden="true">·</span>
                                    <span className="font-mono-tabular">
                                      {new Date(file.modifiedTime).toLocaleString(
                                        'vi-VN'
                                      )}
                                    </span>
                                  </>
                                )}
                              </div>
                            </div>
                          </div>

                          <div className="flex items-center gap-2 shrink-0 self-end sm:self-center">
                            {isBackupJson ? (
                              <button
                                type="button"
                                disabled={isBusy}
                                onClick={() => handleRestoreFromBackupFile(file)}
                                className="inline-flex items-center gap-1 px-2.5 py-1 text-xs font-semibold text-emerald-800 bg-emerald-50 border border-emerald-300 hover:bg-emerald-100 rounded transition-colors whitespace-nowrap shrink-0 cursor-pointer"
                              >
                                <CloudDownload className="w-3.5 h-3.5" />
                                <span>
                                  {isBusy
                                    ? 'Đang nạp...'
                                    : 'Nạp kho mẫu này'}
                                </span>
                              </button>
                            ) : (
                              <button
                                type="button"
                                disabled={isBusy}
                                onClick={() => handleSelectDocxFromDrive(file)}
                                className="inline-flex items-center gap-1 px-2.5 py-1 text-xs font-semibold text-blue-700 bg-blue-50 hover:bg-blue-100 rounded transition-colors whitespace-nowrap shrink-0 cursor-pointer"
                              >
                                <CloudDownload className="w-3.5 h-3.5" />
                                <span>
                                  {isBusy
                                    ? 'Đang tải...'
                                    : 'Chọn làm mẫu Word'}
                                </span>
                              </button>
                            )}

                            {file.webViewLink && (
                              <a
                                href={file.webViewLink}
                                target="_blank"
                                rel="noopener noreferrer"
                                title="Mở trên Google Drive"
                                className="p-1.5 text-slate-500 hover:text-slate-900 rounded hover:bg-slate-100"
                              >
                                <ExternalLink className="w-3.5 h-3.5" />
                              </a>
                            )}

                            <button
                              type="button"
                              disabled={isBusy}
                              onClick={() =>
                                setPendingConfirm({
                                  type: 'delete_file',
                                  file,
                                  title: `Xóa tệp "${file.name}" khỏi Google Drive?`,
                                  description: `Bạn có chắc chắn muốn xóa vĩnh viễn tệp "${file.name}" khỏi Google Drive của bạn không? Hành động này không thể hoàn tác.`,
                                })
                              }
                              title="Xóa tệp khỏi Google Drive"
                              className="p-1.5 text-red-600 hover:text-red-700 rounded hover:bg-red-50 cursor-pointer"
                            >
                              <Trash2 className="w-3.5 h-3.5" />
                            </button>
                          </div>
                        </div>
                      );
                    })}
                  </div>
                )}
              </div>
            </>
          )}

          {/* Connect via Public Google Drive Link / File ID (No Login Required for Visitors) */}
          <div className="p-4 bg-slate-50 border border-slate-200 rounded-lg space-y-2.5">
            <div className="flex items-center gap-2">
              <Link2 className="w-4 h-4 text-blue-700 shrink-0" />
              <h4 className="text-xs font-semibold text-slate-900">
                Liên kết trực tiếp với tệp Kho Mẫu trên Google Drive bằng Link chia sẻ (Không cần đăng nhập)
              </h4>
            </div>
            <p className="text-[11px] text-slate-600">
              Nếu bạn đã có sẵn tệp <code className="font-mono-tabular">templates_store_backup.json</code> trên Google Drive ở chế độ &ldquo;Bất kỳ ai có đường liên kết đều có thể xem&rdquo;, hãy dán link hoặc File ID vào đây để toàn bộ người dùng truy cập ứng dụng lấy dữ liệu từ tệp Drive đó:
            </p>
            <div className="flex flex-col sm:flex-row items-stretch sm:items-center gap-2">
              <input
                type="text"
                value={publicDriveLinkInput}
                onChange={(e) => setPublicDriveLinkInput(e.target.value)}
                placeholder="Dán link Google Drive (https://drive.google.com/file/d/.../view?usp=sharing) hoặc File ID..."
                className="flex-1 px-3 py-2 text-xs bg-white border border-slate-300 rounded-lg text-slate-900 focus:outline-none focus:border-blue-600"
              />
              <button
                type="button"
                disabled={isLinkingPublicDrive || !publicDriveLinkInput.trim()}
                onClick={handleLinkPublicDriveUrl}
                className="inline-flex items-center justify-center gap-1.5 px-3.5 py-2 text-xs font-semibold text-white bg-slate-900 hover:bg-slate-800 disabled:opacity-50 rounded-lg whitespace-nowrap shrink-0 cursor-pointer"
              >
                <FolderSync className="w-3.5 h-3.5" />
                <span>
                  {isLinkingPublicDrive
                    ? 'Đang nạp từ Drive...'
                    : 'Kết nối tệp Drive này'}
                </span>
              </button>
            </div>

            <div className="pt-2 border-t border-slate-200 flex flex-wrap items-center justify-between gap-2">
              <span className="text-[11px] text-slate-500">
                Chạy trên GitHub Pages hoặc không qua proxy? Bạn có thể nạp trực tiếp từ tệp JSON:
              </span>
              <div className="flex flex-wrap items-center gap-2">
                <label className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold text-blue-700 bg-blue-50 border border-blue-200 hover:bg-blue-100 rounded-md cursor-pointer">
                  <CloudUpload className="w-3.5 h-3.5" />
                  <span>Chọn tệp .json từ máy</span>
                  <input
                    type="file"
                    accept=".json,application/json"
                    onChange={handleLocalBackupJsonUpload}
                    className="hidden"
                  />
                </label>
                <button
                  type="button"
                  disabled={isLinkingPublicDrive}
                  onClick={handleLoadStaticBundledLibrary}
                  className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold text-emerald-800 bg-emerald-50 border border-emerald-200 hover:bg-emerald-100 rounded-md cursor-pointer"
                >
                  <CloudDownload className="w-3.5 h-3.5" />
                  <span>Nạp Kho Mẫu Mặc Định (6 mẫu)</span>
                </button>
              </div>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
};
