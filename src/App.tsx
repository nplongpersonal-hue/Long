import React, { useEffect, useMemo, useRef, useState } from 'react';
import {
  AlertTriangle,
  ArrowRightLeft,
  ArrowUpDown,
  Braces,
  Check,
  Cloud,
  CloudUpload,
  Copy,
  CopyPlus,
  Download,
  ExternalLink,
  FileSpreadsheet,
  FileText,
  Filter,
  FolderOpen,
  Globe,
  History,
  MousePointerClick,
  Pencil,
  Plus,
  RefreshCw,
  RotateCcw,
  Search,
  Settings2,
  Share2,
  SlidersHorizontal,
  Tag,
  Trash2,
  Upload,
  X,
} from 'lucide-react';
import { User } from 'firebase/auth';
import {
  arrayBufferToBase64,
  DocxParagraphItem,
  exportFilledDocxBlob,
  parseDocxFromBase64,
  triggerBlobDownload,
} from './utils/docxEngine';
import {
  addExportHistory,
  buildShareableTemplateUrl,
  clearExportHistory,
  convertTemplateTokensToSymbol,
  DEFAULT_FIELD_SYMBOL,
  ensureStarterTemplatesSeeded,
  ExportHistoryRecord,
  extractDriveIdFromUrlQuery,
  extractMergeFieldKeysFromText,
  extractSharedTemplateFromUrlHash,
  fetchSharedDriveConfig,
  FieldSymbolPair,
  formatFieldToken,
  formatPythonCompatibleConfig,
  generateExtraTestUploadDocx,
  getFieldDefByKey,
  loadCustomMailingsFields,
  loadExportHistory,
  loadFieldSymbolConfig,
  loadPublicDriveLibraryByFileId,
  MailingsFieldDef,
  normalizeFieldKeyInput,
  PRESET_FIELD_SYMBOLS,
  removeFieldKeyFromTemplateString,
  replaceFieldKeyInTemplateString,
  resetDefaultMailingsFields,
  resolveParagraphTemplate,
  saveCustomMailingsFields,
  saveFieldSymbolConfig,
  saveTemplatesConfig,
  SharedDrivePublicationInfo,
  syncHistoryWithServer,
  syncTemplatesWithServer,
  TemplatesConfigMap,
} from './utils/templateStore';
import {
  DriveFileItem,
  getAccessToken,
  googleSignIn,
  initAuth,
  logoutGoogle,
  publishLibraryToSharedGoogleDrive,
  SHARED_BACKUP_FILE_NAME,
  SHARED_DRIVE_FOLDER_NAME,
  uploadOrUpdateFileToDrive,
} from './utils/googleDriveService';
import { WordPreviewSheet } from './components/WordPreviewSheet';
import { MailingsPositionEditor } from './components/MailingsPositionEditor';
import { GoogleDriveModal, GoogleSignInButton } from './components/GoogleDriveModal';

type ActiveSection = 'use' | 'configure' | 'library' | 'json';
type LibrarySortOption = 'updated_desc' | 'name_asc' | 'fields_desc' | 'lines_desc';
type HistoryTimeFilter = 'all' | 'today' | '7days' | '30days';

function toSafeDocxFileName(name: string): string {
  return (
    name
      .normalize('NFD')
      .replace(/[\u0300-\u036f]/g, '')
      .replace(/đ/g, 'd')
      .replace(/Đ/g, 'D')
      .toLowerCase()
      .replace(/[^a-z0-9]+/g, '_')
      .replace(/^_+|_+$/g, '') + '.docx'
  );
}

export default function App() {
  const [activeSection, setActiveSection] = useState<ActiveSection>('library');
  const [templatesConfig, setTemplatesConfig] = useState<TemplatesConfigMap>({});
  const [customMailingsFields, setCustomMailingsFields] = useState<MailingsFieldDef[]>([]);
  const [isInitializing, setIsInitializing] = useState<boolean>(true);

  // --- Google Drive OAuth & Shared Drive Publication State ---
  const [googleUser, setGoogleUser] = useState<User | null>(null);
  const [needsAuth, setNeedsAuth] = useState<boolean>(true);
  const [isLoggingIn, setIsLoggingIn] = useState<boolean>(false);
  const [isDriveModalOpen, setIsDriveModalOpen] = useState<boolean>(false);
  const [isSavingToDrive, setIsSavingToDrive] = useState<boolean>(false);
  const [lastDriveSavedFile, setLastDriveSavedFile] = useState<DriveFileItem | null>(null);
  const [sharedDriveInfo, setSharedDriveInfo] = useState<SharedDrivePublicationInfo | null>(null);
  const [isConfirmingDrivePublish, setIsConfirmingDrivePublish] = useState<boolean>(false);
  const [isPublishingSharedDrive, setIsPublishingSharedDrive] = useState<boolean>(false);

  // --- Tab 1: Use Saved Template (Mailings Fill & Line Fill) State ---
  const [chosenTemplate, setChosenTemplate] = useState<string>('');
  const [parsedTemplateParagraphs, setParsedTemplateParagraphs] = useState<DocxParagraphItem[]>([]);
  const [fillMode, setFillMode] = useState<'mailings_fields' | 'full_lines' | 'all_document_lines'>('mailings_fields');
  const [useSectionFilter, setUseSectionFilter] = useState<'all' | 'header' | 'body' | 'footnote' | 'footer'>('all');
  const [mailingsFieldValues, setMailingsFieldValues] = useState<Record<string, string>>({});
  const [userReplacements, setUserReplacements] = useState<Record<number, string>>({});
  const [activeFieldIndex, setActiveFieldIndex] = useState<number | null>(null);
  const [activeFieldKey, setActiveFieldKey] = useState<string | null>(null);
  const [showLineNumbersPreview, setShowLineNumbersPreview] = useState<boolean>(true);
  const [editingTab1FieldKey, setEditingTab1FieldKey] = useState<string | null>(null);
  const [tab1RenameLabel, setTab1RenameLabel] = useState<string>('');
  const [tab1RenameKey, setTab1RenameKey] = useState<string>('');
  const [showTab1SymbolPicker, setShowTab1SymbolPicker] = useState<boolean>(false);
  const [tab1CustomOpen, setTab1CustomOpen] = useState<string>('«');
  const [tab1CustomClose, setTab1CustomClose] = useState<string>('»');
  const [lastExportedBlob, setLastExportedBlob] = useState<{
    blob: Blob;
    fileName: string;
    templateName: string;
  } | null>(null);
  const [isExporting, setIsExporting] = useState<boolean>(false);
  const [useStatusMessage, setUseStatusMessage] = useState<{
    type: 'success' | 'warning' | 'error';
    text: string;
  } | null>(null);

  // --- Tab 2: Add / Configure Template (Mailings Word Position Configurator) State ---
  const [fieldSymbol, setFieldSymbol] = useState<FieldSymbolPair>(() => loadFieldSymbolConfig());
  const [templateNameInput, setTemplateNameInput] = useState<string>('');
  const [templateCategoryInput, setTemplateCategoryInput] = useState<string>('Hành chính - Tổng hợp');
  const [uploadedFileName, setUploadedFileName] = useState<string>('');
  const [uploadedBase64, setUploadedBase64] = useState<string>('');
  const [uploadedAllParagraphs, setUploadedAllParagraphs] = useState<DocxParagraphItem[]>([]);
  const [uploadedNonEmptyParagraphs, setUploadedNonEmptyParagraphs] = useState<DocxParagraphItem[]>([]);
  const [chosenIndices, setChosenIndices] = useState<string[]>([]);
  const [paragraphTemplatesInput, setParagraphTemplatesInput] = useState<Record<string, string>>({});
  const [detectedDefaultValues, setDetectedDefaultValues] = useState<Record<string, string>>({});
  const [customLabelsInput, setCustomLabelsInput] = useState<Record<string, string>>({});
  const [activeConfigParagraphIndex, setActiveConfigParagraphIndex] = useState<number | null>(null);
  const [paragraphSearchQuery, setParagraphSearchQuery] = useState<string>('');
  const [configureSectionFilter, setConfigureSectionFilter] = useState<'all' | 'header' | 'body' | 'footnote' | 'footer'>('all');
  const [configStatusMessage, setConfigStatusMessage] = useState<{
    type: 'success' | 'warning' | 'error';
    text: string;
  } | null>(null);
  const [editingExistingTemplateName, setEditingExistingTemplateName] = useState<string | null>(null);

  // --- Duplicate (Clone) Template Dialog / Panel State ---
  const [duplicatingSourceName, setDuplicatingSourceName] = useState<string | null>(null);
  const [duplicateNewName, setDuplicateNewName] = useState<string>('');
  const [duplicateNewCategory, setDuplicateNewCategory] = useState<string>('');
  const [openConfigAfterDuplicate, setOpenConfigAfterDuplicate] = useState<boolean>(false);
  const [duplicateError, setDuplicateError] = useState<string | null>(null);
  const [libraryBannerMessage, setLibraryBannerMessage] = useState<string | null>(null);

  // --- Tab 3: Library & History Search / Filter State ---
  const [librarySearchQuery, setLibrarySearchQuery] = useState<string>('');
  const [libraryCategoryFilter, setLibraryCategoryFilter] = useState<string>('ALL');
  const [librarySortBy, setLibrarySortBy] = useState<LibrarySortOption>('updated_desc');

  const [exportHistory, setExportHistory] = useState<ExportHistoryRecord[]>([]);
  const [historySearchQuery, setHistorySearchQuery] = useState<string>('');
  const [historyTemplateFilter, setHistoryTemplateFilter] = useState<string>('ALL');
  const [historyTimeFilter, setHistoryTimeFilter] = useState<HistoryTimeFilter>('all');

  // --- Tab 4: JSON State ---
  const [copiedJson, setCopiedJson] = useState<boolean>(false);
  const [jsonImportText, setJsonImportText] = useState<string>('');
  const [jsonStatusMessage, setJsonStatusMessage] = useState<string | null>(null);

  // Refs for focusing inputs when clicking preview lines or «Field» tags
  const fieldInputRefs = useRef<Record<number, HTMLInputElement | HTMLTextAreaElement | null>>({});
  const mailingsInputRefs = useRef<Record<string, HTMLInputElement | null>>({});
  const fileInputRef = useRef<HTMLInputElement | null>(null);
  const configureLeftPanelRef = useRef<HTMLDivElement | null>(null);
  const mailingsEditorAnchorRef = useRef<HTMLDivElement | null>(null);

  // Initialize Firebase Google Auth listener
  useEffect(() => {
    const unsubscribe = initAuth(
      (user) => {
        setGoogleUser(user);
        setNeedsAuth(false);
      },
      () => {
        setGoogleUser(null);
        setNeedsAuth(true);
      }
    );
    return () => unsubscribe();
  }, []);

  const handleGoogleLogin = async () => {
    setIsLoggingIn(true);
    try {
      const result = await googleSignIn();
      if (result) {
        setGoogleUser(result.user);
        setNeedsAuth(false);
      }
    } catch (err) {
      console.error('Google Sign-In failed:', err);
    } finally {
      setIsLoggingIn(false);
    }
  };

  const handleGoogleLogout = async () => {
    await logoutGoogle();
    setGoogleUser(null);
    setNeedsAuth(true);
  };

  // Initialize starter templates & sync with Shared Google Drive / server store on first load
  useEffect(() => {
    let mounted = true;
    async function init() {
      setIsInitializing(true);
      try {
        const sharedFromHash = extractSharedTemplateFromUrlHash();
        const driveIdFromUrl = extractDriveIdFromUrlQuery();

        let driveLoadedConfig: TemplatesConfigMap | null = null;
        let activeDriveInfo: SharedDrivePublicationInfo | null = null;

        if (driveIdFromUrl) {
          try {
            const linked = await loadPublicDriveLibraryByFileId(driveIdFromUrl);
            driveLoadedConfig = linked.templates;
            activeDriveInfo = linked.driveConfig;
          } catch (err) {
            console.warn('Could not load ?driveId from Google Drive:', err);
          }
        }

        if (!driveLoadedConfig) {
          const driveStatus = await fetchSharedDriveConfig(true);
          if (driveStatus.driveConfig) {
            activeDriveInfo = driveStatus.driveConfig;
          }
          if (Object.keys(driveStatus.templates).length > 0) {
            driveLoadedConfig = driveStatus.templates;
          }
        }

        const seeded =
          driveLoadedConfig && Object.keys(driveLoadedConfig).length > 0
            ? driveLoadedConfig
            : await ensureStarterTemplatesSeeded();

        const sharedHistory = await syncHistoryWithServer();
        if (!mounted) return;

        let finalConfig = seeded;
        if (sharedFromHash) {
          finalConfig = {
            ...seeded,
            [sharedFromHash.name]: sharedFromHash.entry,
          };
          saveTemplatesConfig(finalConfig);
        }

        setSharedDriveInfo(activeDriveInfo);
        setTemplatesConfig(finalConfig);
        setCustomMailingsFields(loadCustomMailingsFields());
        setExportHistory(sharedHistory.length > 0 ? sharedHistory : loadExportHistory());

        if (sharedFromHash) {
          setChosenTemplate(sharedFromHash.name);
          setActiveSection('use');
          setUseStatusMessage({
            type: 'success',
            text: `Đã mở và lưu biểu mẫu được chia sẻ "${sharedFromHash.name}" vào Kho Biểu Mẫu chung!`,
          });
          setLibraryBannerMessage(
            `Đã nhận biểu mẫu được chia sẻ "${sharedFromHash.name}" và đồng bộ vào kho mẫu!`
          );
        } else {
          const names = Object.keys(finalConfig);
          if (names.length > 0) {
            setChosenTemplate(names[0]);
          }
        }
      } finally {
        if (mounted) setIsInitializing(false);
      }
    }
    init();
    return () => {
      mounted = false;
    };
  }, []);

  // Trigger confirmation dialog to publish & share all templates on the user's Google Drive
  const handleRequestPublishToSharedDrive = async () => {
    const token = await getAccessToken();
    if (!token || needsAuth || !googleUser) {
      setIsDriveModalOpen(true);
      return;
    }
    setIsConfirmingDrivePublish(true);
  };

  // Confirmed handler: uploads all templates & .docx files to Google Drive and sets 'anyone' reader permission
  const handleConfirmPublishToSharedDrive = async () => {
    setIsConfirmingDrivePublish(false);
    setIsPublishingSharedDrive(true);
    try {
      const { driveInfo, updatedTemplatesConfig } =
        await publishLibraryToSharedGoogleDrive({
          templatesConfig,
          ownerEmail: googleUser?.email || undefined,
          existingBackupFileId: sharedDriveInfo?.backupFileId,
          uploadIndividualDocxFiles: true,
        });

      setSharedDriveInfo(driveInfo);
      setTemplatesConfig(updatedTemplatesConfig);
      saveTemplatesConfig(updatedTemplatesConfig);

      const appDriveShareUrl = `${window.location.origin}${window.location.pathname}?driveId=${encodeURIComponent(
        driveInfo.backupFileId
      )}`;
      try {
        await navigator.clipboard.writeText(appDriveShareUrl);
      } catch {
        // Ignore clipboard error
      }

      setLibraryBannerMessage(
        `Đã lưu toàn bộ ${driveInfo.templateCount} biểu mẫu vào thư mục "${SHARED_DRIVE_FOLDER_NAME}" trên Google Drive (${
          googleUser?.email || ''
        }) và bật quyền chia sẻ công khai! Đã sao chép Link App kết nối trực tiếp với Google Drive của bạn.`
      );
    } catch (err: any) {
      if (err?.message === 'UNAUTHENTICATED') {
        setNeedsAuth(true);
        setIsDriveModalOpen(true);
      } else {
        setLibraryBannerMessage(
          err instanceof Error
            ? err.message
            : 'Không thể lưu và chia sẻ kho mẫu lên Google Drive.'
        );
      }
    } finally {
      setIsPublishingSharedDrive(false);
    }
  };

  // Share the entire template library link (backed by Google Drive if published)
  const handleShareLibraryLink = async () => {
    saveTemplatesConfig(templatesConfig);
    await syncTemplatesWithServer();
    const shareUrl = sharedDriveInfo?.backupFileId
      ? `${window.location.origin}${window.location.pathname}?driveId=${encodeURIComponent(
          sharedDriveInfo.backupFileId
        )}`
      : `${window.location.origin}${window.location.pathname}`;
    try {
      await navigator.clipboard.writeText(shareUrl);
      setLibraryBannerMessage(
        sharedDriveInfo
          ? `Đã sao chép Link App kết nối trực tiếp với Google Drive (${shareUrl})! Bất kỳ ai mở link này đều tự động tải ${Object.keys(templatesConfig).length} mẫu từ Google Drive của bạn.`
          : `Đã sao chép Link Kho Biểu Mẫu (${shareUrl})! Để lưu trữ trên Google Drive và chia sẻ cho mọi người, hãy nhấn "Lưu & Chia Sẻ Lên Google Drive".`
      );
    } catch {
      setLibraryBannerMessage(`Link chia sẻ Kho Biểu Mẫu: ${shareUrl}`);
    }
  };

  // Share a direct link to a specific template (both synced on server and encoded in URL hash)
  const handleShareSingleTemplateLink = async (templateName: string) => {
    const entry = templatesConfig[templateName];
    if (!entry) return;
    saveTemplatesConfig(templatesConfig);
    const url = buildShareableTemplateUrl(templateName, entry);
    try {
      await navigator.clipboard.writeText(url);
      setLibraryBannerMessage(
        `Đã sao chép link chia sẻ trực tiếp cho mẫu "${templateName}"! Người nhận mở link sẽ có sẵn mẫu này để điền ngay.`
      );
      setUseStatusMessage({
        type: 'success',
        text: `Đã sao chép link chia sẻ cho mẫu "${templateName}" vào bộ nhớ tạm!`,
      });
    } catch {
      setLibraryBannerMessage(`Link chia sẻ mẫu "${templateName}": ${url}`);
    }
  };

  // Refresh shared templates directly from Google Drive (and server fallback)
  const handleRefreshSharedLibrary = async () => {
    const driveStatus = await fetchSharedDriveConfig(true);
    if (driveStatus.driveConfig) {
      setSharedDriveInfo(driveStatus.driveConfig);
    }
    const latestTemplates =
      Object.keys(driveStatus.templates).length > 0
        ? driveStatus.templates
        : await syncTemplatesWithServer();
    const latestHistory = await syncHistoryWithServer();
    setTemplatesConfig(latestTemplates);
    setExportHistory(latestHistory);
    setLibraryBannerMessage(
      driveStatus.syncedFromDrive
        ? `Đã tải trực tiếp ${Object.keys(latestTemplates).length} biểu mẫu mới nhất từ Google Drive dùng chung!`
        : `Đã đồng bộ Kho Biểu Mẫu Chung mới nhất (${Object.keys(latestTemplates).length} biểu mẫu)!`
    );
  };

  const templateNames = useMemo(() => Object.keys(templatesConfig), [templatesConfig]);
  const currentTemplateInfo = chosenTemplate ? templatesConfig[chosenTemplate] : undefined;

  // Unique categories across all saved templates for filtering in Kho Mẫu
  const availableCategories = useMemo(() => {
    const cats = new Set<string>();
    Object.values(templatesConfig).forEach((entry) => {
      cats.add(entry.category || 'Biểu mẫu chung');
    });
    return Array.from(cats);
  }, [templatesConfig]);

  // Filtered and sorted templates in Kho Biểu Mẫu (Tab 3)
  const filteredLibraryTemplates = useMemo(() => {
    const q = librarySearchQuery.trim().toLowerCase();
    const list = Object.entries(templatesConfig).filter(([name, entry]) => {
      const cat = entry.category || 'Biểu mẫu chung';
      if (libraryCategoryFilter !== 'ALL' && cat !== libraryCategoryFilter) {
        return false;
      }
      if (!q) return true;

      const fieldKeys = entry.mailings_fields ? Object.keys(entry.mailings_fields) : [];
      const fieldLabels = entry.mailings_fields
        ? Object.values(entry.mailings_fields).map((m) => m.label)
        : [];

      const matchName = name.toLowerCase().includes(q);
      const matchCat = cat.toLowerCase().includes(q);
      const matchPath = entry.file_path.toLowerCase().includes(q);
      const matchFields =
        fieldKeys.some((k) => k.toLowerCase().includes(q)) ||
        fieldLabels.some((l) => l.toLowerCase().includes(q));

      return matchName || matchCat || matchPath || matchFields;
    });

    list.sort(([nameA, entryA], [nameB, entryB]) => {
      if (librarySortBy === 'name_asc') {
        return nameA.localeCompare(nameB, 'vi');
      }
      if (librarySortBy === 'fields_desc') {
        const countA = entryA.mailings_fields ? Object.keys(entryA.mailings_fields).length : 0;
        const countB = entryB.mailings_fields ? Object.keys(entryB.mailings_fields).length : 0;
        return countB - countA;
      }
      if (librarySortBy === 'lines_desc') {
        return entryB.selected_indices.length - entryA.selected_indices.length;
      }
      return (entryB.updated_at || '').localeCompare(entryA.updated_at || '');
    });

    return list;
  }, [templatesConfig, librarySearchQuery, libraryCategoryFilter, librarySortBy]);

  // Unique template names present in export history for filter dropdown
  const historyTemplateOptions = useMemo(() => {
    const set = new Set<string>();
    exportHistory.forEach((rec) => set.add(rec.templateName));
    return Array.from(set);
  }, [exportHistory]);

  // Filtered export history records in Tab 3
  const filteredExportHistory = useMemo(() => {
    const q = historySearchQuery.trim().toLowerCase();
    const now = Date.now();
    const startOfToday = new Date();
    startOfToday.setHours(0, 0, 0, 0);

    return exportHistory.filter((rec) => {
      if (historyTemplateFilter !== 'ALL' && rec.templateName !== historyTemplateFilter) {
        return false;
      }

      if (historyTimeFilter !== 'all') {
        const exportedTime = new Date(rec.exportedAt).getTime();
        if (historyTimeFilter === 'today' && exportedTime < startOfToday.getTime()) {
          return false;
        }
        if (historyTimeFilter === '7days' && now - exportedTime > 7 * 24 * 3600 * 1000) {
          return false;
        }
        if (historyTimeFilter === '30days' && now - exportedTime > 30 * 24 * 3600 * 1000) {
          return false;
        }
      }

      if (!q) return true;

      const matchOutput = rec.outputFileName.toLowerCase().includes(q);
      const matchTpl = rec.templateName.toLowerCase().includes(q);
      const matchFieldValues =
        rec.fieldValues &&
        Object.values(rec.fieldValues).some((v) => v.toLowerCase().includes(q));
      const matchReplacements =
        rec.replacements &&
        Object.values(rec.replacements).some((v) => v.toLowerCase().includes(q));

      return Boolean(matchOutput || matchTpl || matchFieldValues || matchReplacements);
    });
  }, [exportHistory, historySearchQuery, historyTemplateFilter, historyTimeFilter]);

  // Extract ordered list of Mailings fields used in the currently selected template (Tab 1)
  const activeTemplateMailingsFields = useMemo(() => {
    if (!currentTemplateInfo) return [];
    const pTemplates = currentTemplateInfo.paragraph_templates || {};
    const fieldMap = new Map<
      string,
      {
        key: string;
        label: string;
        placeholder: string;
        defaultValue: string;
        lineIndices: number[];
      }
    >();

    currentTemplateInfo.selected_indices.forEach((idxStr) => {
      const idx = Number(idxStr);
      const tpl = pTemplates[idxStr];
      if (!tpl) return;
      const keys = extractMergeFieldKeysFromText(tpl);
      keys.forEach((k) => {
        const def = getFieldDefByKey(k, customMailingsFields);
        const savedMeta = currentTemplateInfo.mailings_fields?.[k];
        if (!fieldMap.has(k)) {
          fieldMap.set(k, {
            key: k,
            label: savedMeta?.label || def.label,
            placeholder: def.placeholder,
            defaultValue: savedMeta?.defaultValue ?? '',
            lineIndices: [idx],
          });
        } else {
          const existing = fieldMap.get(k)!;
          if (!existing.lineIndices.includes(idx)) {
            existing.lineIndices.push(idx);
          }
        }
      });
    });

    return Array.from(fieldMap.values());
  }, [currentTemplateInfo, customMailingsFields]);

  // Parse the currently selected template in Tab 1 whenever `chosenTemplate` or `templatesConfig` changes
  useEffect(() => {
    let cancelled = false;
    async function loadSelectedTemplateDocx() {
      if (!chosenTemplate || !templatesConfig[chosenTemplate]) {
        setParsedTemplateParagraphs([]);
        setUserReplacements({});
        setMailingsFieldValues({});
        setLastExportedBlob(null);
        setLastDriveSavedFile(null);
        return;
      }

      const info = templatesConfig[chosenTemplate];
      try {
        const parsed = await parseDocxFromBase64(info.base64_docx);
        if (cancelled) return;
        setParsedTemplateParagraphs(parsed.allParagraphs);

        const initFieldVals: Record<string, string> = {};
        if (info.mailings_fields) {
          Object.entries(info.mailings_fields).forEach(([k, meta]) => {
            initFieldVals[k] = meta.defaultValue || '';
          });
        }
        if (info.paragraph_templates) {
          Object.values(info.paragraph_templates).forEach((tpl) => {
            extractMergeFieldKeysFromText(tpl).forEach((k) => {
              if (initFieldVals[k] === undefined) {
                initFieldVals[k] = getFieldDefByKey(k, customMailingsFields).placeholder || '';
              }
            });
          });
        }
        setMailingsFieldValues(initFieldVals);

        const initialMap: Record<number, string> = {};
        info.selected_indices.forEach((idxStr) => {
          const idx = Number(idxStr);
          const original =
            idx >= 0 && idx < parsed.allParagraphs.length
              ? parsed.allParagraphs[idx].text
              : 'Dòng không tồn tại';
          const pTpl = info.paragraph_templates?.[idxStr];
          initialMap[idx] = pTpl ? resolveParagraphTemplate(pTpl, initFieldVals) : original;
        });

        setUserReplacements(initialMap);
        const hasMailings =
          info.paragraph_templates &&
          Object.values(info.paragraph_templates).some(
            (t) => extractMergeFieldKeysFromText(t).length > 0
          );
        setFillMode(hasMailings ? 'mailings_fields' : 'full_lines');
        setLastExportedBlob(null);
        setLastDriveSavedFile(null);
        setUseStatusMessage(null);
      } catch (err) {
        if (!cancelled) {
          setUseStatusMessage({
            type: 'error',
            text: err instanceof Error ? err.message : 'Không thể đọc cấu trúc file Word mẫu.',
          });
        }
      }
    }
    loadSelectedTemplateDocx();
    return () => {
      cancelled = true;
    };
  }, [chosenTemplate, templatesConfig, customMailingsFields]);

  // Whenever user updates a Mailings field value in Tab 1, automatically resolve all paragraphs that contain «Field_Key»
  const handleUpdateMailingsFieldValue = (fieldKey: string, newVal: string) => {
    const nextFieldValues = {
      ...mailingsFieldValues,
      [fieldKey]: newVal,
    };
    setMailingsFieldValues(nextFieldValues);

    if (!currentTemplateInfo) return;
    const pTemplates = currentTemplateInfo.paragraph_templates || {};
    setUserReplacements((prev) => {
      const nextReplacements = { ...prev };
      Object.entries(pTemplates).forEach(([idxStr, tpl]) => {
        const idx = Number(idxStr);
        if (tpl && extractMergeFieldKeysFromText(tpl).length > 0) {
          nextReplacements[idx] = resolveParagraphTemplate(tpl, nextFieldValues);
        }
      });
      return nextReplacements;
    });
  };

  const modifiedFieldsCount = useMemo(() => {
    if (!currentTemplateInfo || parsedTemplateParagraphs.length === 0) return 0;
    let count = 0;
    parsedTemplateParagraphs.forEach((p) => {
      if (userReplacements[p.index] !== undefined && userReplacements[p.index] !== p.text) {
        count++;
      }
    });
    return count;
  }, [currentTemplateInfo, parsedTemplateParagraphs, userReplacements]);

  // --- Duplicate (Clone) Template Handlers ---
  const handleOpenDuplicateModal = (sourceName: string) => {
    const sourceEntry = templatesConfig[sourceName];
    if (!sourceEntry) return;

    let suggested = `${sourceName} (Bản sao)`;
    let counter = 2;
    while (templatesConfig[suggested]) {
      suggested = `${sourceName} (Bản sao ${counter})`;
      counter++;
    }

    setDuplicatingSourceName(sourceName);
    setDuplicateNewName(suggested);
    setDuplicateNewCategory(sourceEntry.category || 'Biểu mẫu chung');
    setOpenConfigAfterDuplicate(false);
    setDuplicateError(null);
  };

  const handleConfirmDuplicateTemplate = async () => {
    if (!duplicatingSourceName) return;
    const sourceEntry = templatesConfig[duplicatingSourceName];
    if (!sourceEntry) return;

    const trimmedNewName = duplicateNewName.trim();
    if (!trimmedNewName) {
      setDuplicateError('Vui lòng nhập tên cho biểu mẫu nhân bản.');
      return;
    }
    if (templatesConfig[trimmedNewName]) {
      setDuplicateError(`Tên mẫu "${trimmedNewName}" đã tồn tại. Vui lòng chọn tên khác.`);
      return;
    }

    const safeFileName = toSafeDocxFileName(trimmedNewName);
    const clonedEntry: TemplatesConfigMap[string] = {
      file_path: `templates_store/${safeFileName}`,
      selected_indices: [...sourceEntry.selected_indices],
      paragraph_templates: sourceEntry.paragraph_templates
        ? { ...sourceEntry.paragraph_templates }
        : undefined,
      mailings_fields: sourceEntry.mailings_fields
        ? JSON.parse(JSON.stringify(sourceEntry.mailings_fields))
        : undefined,
      custom_labels: sourceEntry.custom_labels
        ? { ...sourceEntry.custom_labels }
        : undefined,
      base64_docx: sourceEntry.base64_docx,
      updated_at: new Date().toISOString(),
      total_paragraphs: sourceEntry.total_paragraphs,
      category: duplicateNewCategory.trim() || sourceEntry.category || 'Biểu mẫu chung',
    };

    const nextConfig: TemplatesConfigMap = {
      ...templatesConfig,
      [trimmedNewName]: clonedEntry,
    };

    saveTemplatesConfig(nextConfig);
    setTemplatesConfig(nextConfig);
    setChosenTemplate(trimmedNewName);
    setDuplicatingSourceName(null);

    if (openConfigAfterDuplicate) {
      const parsed = await parseDocxFromBase64(clonedEntry.base64_docx);
      setTemplateNameInput(trimmedNewName);
      setTemplateCategoryInput(clonedEntry.category || 'Biểu mẫu chung');
      setUploadedFileName(safeFileName);
      setUploadedBase64(clonedEntry.base64_docx);
      setUploadedAllParagraphs(parsed.allParagraphs);
      setUploadedNonEmptyParagraphs(parsed.nonEmptyParagraphs);
      setChosenIndices([...clonedEntry.selected_indices]);
      setParagraphTemplatesInput({ ...(clonedEntry.paragraph_templates || {}) });
      const defaults: Record<string, string> = {};
      if (clonedEntry.mailings_fields) {
        Object.entries(clonedEntry.mailings_fields).forEach(([k, v]) => {
          defaults[k] = v.defaultValue;
        });
      }
      setDetectedDefaultValues(defaults);
      setCustomLabelsInput({ ...(clonedEntry.custom_labels || {}) });
      setEditingExistingTemplateName(trimmedNewName);
      const firstIdx =
        clonedEntry.selected_indices.length > 0
          ? Number(clonedEntry.selected_indices[0])
          : parsed.nonEmptyParagraphs[0]?.index ?? null;
      setActiveConfigParagraphIndex(firstIdx);
      setConfigStatusMessage({
        type: 'success',
        text: `Đã nhân bản thành công mẫu "${trimmedNewName}". Bạn có thể tùy chỉnh vị trí trường cho bản sao này.`,
      });
      setActiveSection('configure');
    } else {
      setLibraryBannerMessage(
        `Đã nhân bản thành công mẫu "${duplicatingSourceName}" thành "${trimmedNewName}"!`
      );
      setUseStatusMessage({
        type: 'success',
        text: `Đã nhân bản và chuyển sang mẫu "${trimmedNewName}"!`,
      });
    }
  };

  // --- Tab 2 Handlers: Uploading & Configuring Mailings Positions ---
  const handleFileUpload = async (file: File) => {
    setConfigStatusMessage(null);
    if (!file.name.toLowerCase().endsWith('.docx')) {
      setConfigStatusMessage({
        type: 'error',
        text: 'Vui lòng chọn file định dạng Microsoft Word (.docx).',
      });
      return;
    }

    try {
      const buffer = await file.arrayBuffer();
      const base64 = arrayBufferToBase64(buffer);
      const parsed = await parseDocxFromBase64(base64);

      setUploadedFileName(file.name);
      setUploadedBase64(base64);
      setUploadedAllParagraphs(parsed.allParagraphs);
      setUploadedNonEmptyParagraphs(parsed.nonEmptyParagraphs);

      const cleanName = file.name.replace(/\.docx$/i, '').replace(/[_-]+/g, ' ');
      const capitalized = cleanName.charAt(0).toUpperCase() + cleanName.slice(1);
      setTemplateNameInput(capitalized);

      setChosenIndices([]);
      setParagraphTemplatesInput({});
      setDetectedDefaultValues({});
      setCustomLabelsInput({});
      setEditingExistingTemplateName(null);
      setActiveConfigParagraphIndex(parsed.nonEmptyParagraphs[0]?.index ?? null);

      setConfigStatusMessage({
        type: 'success',
        text: `Đã tải lên "${file.name}" (${parsed.nonEmptyParagraphs.length} dòng văn bản). Hãy nhấp vào dòng cần chỉnh sửa để đặt vị trí trường dữ liệu.`,
      });
    } catch (err) {
      setConfigStatusMessage({
        type: 'error',
        text:
          err instanceof Error
            ? err.message
            : 'Không thể phân tích file .docx. Vui lòng kiểm tra lại định dạng file.',
      });
    }
  };

  const handleLoadTestUploadDocx = async () => {
    const { fileName, blob } = await generateExtraTestUploadDocx();
    const file = new File([blob], fileName, {
      type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    });
    setTemplateNameInput('Giấy đề nghị cấp văn phòng phẩm');
    setTemplateCategoryInput('Hành chính - Quản trị');
    await handleFileUpload(file);
  };

  const handleEditExistingTemplateConfig = async (templateName: string) => {
    const entry = templatesConfig[templateName];
    if (!entry) return;

    try {
      const parsed = await parseDocxFromBase64(entry.base64_docx);
      setTemplateNameInput(templateName);
      setTemplateCategoryInput(entry.category || 'Hành chính - Tổng hợp');
      setUploadedFileName(entry.file_path.split('/').pop() || `${templateName}.docx`);
      setUploadedBase64(entry.base64_docx);
      setUploadedAllParagraphs(parsed.allParagraphs);
      setUploadedNonEmptyParagraphs(parsed.nonEmptyParagraphs);
      setChosenIndices([...entry.selected_indices]);
      setParagraphTemplatesInput({ ...(entry.paragraph_templates || {}) });

      const defaults: Record<string, string> = {};
      if (entry.mailings_fields) {
        Object.entries(entry.mailings_fields).forEach(([k, v]) => {
          defaults[k] = v.defaultValue;
        });
      }
      setDetectedDefaultValues(defaults);
      setCustomLabelsInput({ ...(entry.custom_labels || {}) });
      setEditingExistingTemplateName(templateName);

      const firstIdx =
        entry.selected_indices.length > 0
          ? Number(entry.selected_indices[0])
          : parsed.nonEmptyParagraphs[0]?.index ?? null;
      setActiveConfigParagraphIndex(firstIdx);

      setConfigStatusMessage({
        type: 'success',
        text: `Đang chỉnh sửa vị trí trường Mailings cho mẫu "${templateName}". Nhấp vào dòng bất kỳ để đặt con trỏ và chèn trường.`,
      });
      setActiveSection('configure');
    } catch (err) {
      setConfigStatusMessage({
        type: 'error',
        text: err instanceof Error ? err.message : 'Lỗi khi mở cấu hình mẫu.',
      });
    }
  };

  const handleUpdateParagraphTemplateInConfig = (
    paragraphIndex: number,
    newTemplateStr: string,
    extractedDefaultVal?: { key: string; value: string }
  ) => {
    const idxStr = String(paragraphIndex);
    setParagraphTemplatesInput((prev) => ({
      ...prev,
      [idxStr]: newTemplateStr,
    }));

    setChosenIndices((prev) => {
      if (prev.includes(idxStr)) return prev;
      return [...prev, idxStr].sort((a, b) => Number(a) - Number(b));
    });

    if (extractedDefaultVal && extractedDefaultVal.value) {
      setDetectedDefaultValues((prev) => ({
        ...prev,
        [extractedDefaultVal.key]: extractedDefaultVal.value,
      }));
    }
  };

  const handleResetParagraphTemplateInConfig = (paragraphIndex: number) => {
    const idxStr = String(paragraphIndex);
    setParagraphTemplatesInput((prev) => {
      const next = { ...prev };
      delete next[idxStr];
      return next;
    });
    setChosenIndices((prev) => prev.filter((item) => item !== idxStr));
  };

  const handleToggleConfigureIndex = (idx: number) => {
    const idxStr = String(idx);
    setActiveConfigParagraphIndex(idx);
    setChosenIndices((prev) => {
      if (prev.includes(idxStr)) {
        return prev.filter((item) => item !== idxStr);
      }
      return [...prev, idxStr].sort((a, b) => Number(a) - Number(b));
    });
  };

  const currentTemplateSymbol: FieldSymbolPair = useMemo(
    () => currentTemplateInfo?.field_symbol || fieldSymbol || DEFAULT_FIELD_SYMBOL,
    [currentTemplateInfo, fieldSymbol]
  );

  // Update field delimiter symbol in Tab 1 (Use Template) and persist to template
  const handleUpdateFieldSymbolInTab1 = (newSymbol: FieldSymbolPair) => {
    setFieldSymbol(newSymbol);
    saveFieldSymbolConfig(newSymbol);
    setTab1CustomOpen(newSymbol.openSymbol);
    setTab1CustomClose(newSymbol.closeSymbol);

    if (!chosenTemplate || !currentTemplateInfo) return;
    const oldSymbol = currentTemplateInfo.field_symbol || fieldSymbol;
    const updatedPTemplates: Record<string, string> = {};
    Object.entries(currentTemplateInfo.paragraph_templates || {}).forEach(([idxStr, tpl]) => {
      updatedPTemplates[idxStr] = convertTemplateTokensToSymbol(tpl, newSymbol, oldSymbol);
    });

    const nextConfig: TemplatesConfigMap = {
      ...templatesConfig,
      [chosenTemplate]: {
        ...currentTemplateInfo,
        field_symbol: newSymbol,
        paragraph_templates: updatedPTemplates,
        updated_at: new Date().toISOString(),
      },
    };
    saveTemplatesConfig(nextConfig);
    setTemplatesConfig(nextConfig);
    setUseStatusMessage({
      type: 'success',
      text: `Đã đổi ký hiệu trường sang ${newSymbol.openSymbol}Ten_Truong${newSymbol.closeSymbol} cho mẫu "${chosenTemplate}"!`,
    });
  };

  // Rename field label, change field key, or swap field right in Tab 1 (Use Template)
  const handleRenameOrSwapFieldInTab1 = (
    oldKey: string,
    newLabel: string,
    newKeyRaw: string
  ) => {
    if (!chosenTemplate || !currentTemplateInfo) return;
    const cleanNewKey = normalizeFieldKeyInput(newKeyRaw || newLabel) || oldKey;
    const cleanNewLabel = newLabel.trim() || getFieldDefByKey(cleanNewKey, customMailingsFields).label;
    const activeSym = currentTemplateSymbol;

    // 1. Update paragraph_templates in current template
    const nextPTemplates: Record<string, string> = {};
    Object.entries(currentTemplateInfo.paragraph_templates || {}).forEach(([idxStr, tpl]) => {
      nextPTemplates[idxStr] = replaceFieldKeyInTemplateString(
        tpl,
        oldKey,
        cleanNewKey,
        activeSym
      );
    });

    // 2. Update mailings_fields metadata in current template
    const nextMailingsFields = { ...(currentTemplateInfo.mailings_fields || {}) };
    const oldMeta = nextMailingsFields[oldKey];
    const currentInputVal = mailingsFieldValues[oldKey] ?? oldMeta?.defaultValue ?? '';
    if (oldKey !== cleanNewKey) {
      delete nextMailingsFields[oldKey];
    }
    nextMailingsFields[cleanNewKey] = {
      label: cleanNewLabel,
      defaultValue: currentInputVal,
      category: oldMeta?.category || 'Trường tùy chỉnh',
    };

    // 3. Update mailingsFieldValues in state
    const nextFieldVals = { ...mailingsFieldValues };
    if (oldKey !== cleanNewKey) {
      delete nextFieldVals[oldKey];
    }
    nextFieldVals[cleanNewKey] = currentInputVal;
    setMailingsFieldValues(nextFieldVals);

    // 4. Also update palette so the renamed field is available everywhere
    setCustomMailingsFields((prev) => {
      const existsOld = prev.some((f) => f.key === oldKey);
      const existsNew = prev.some((f) => f.key === cleanNewKey);
      let next: MailingsFieldDef[];
      const newDef: MailingsFieldDef = {
        key: cleanNewKey,
        label: cleanNewLabel,
        category: oldMeta?.category || 'Trường tùy chỉnh',
        placeholder: currentInputVal || `Nhập ${cleanNewLabel.toLowerCase()}...`,
      };
      if (existsOld && oldKey !== cleanNewKey) {
        next = prev
          .filter((f) => f.key !== cleanNewKey)
          .map((f) => (f.key === oldKey ? newDef : f));
      } else if (existsNew) {
        next = prev.map((f) => (f.key === cleanNewKey ? newDef : f));
      } else {
        next = [...prev, newDef];
      }
      saveCustomMailingsFields(next);
      return next;
    });

    // 5. Save updated template entry
    const nextConfig: TemplatesConfigMap = {
      ...templatesConfig,
      [chosenTemplate]: {
        ...currentTemplateInfo,
        paragraph_templates: nextPTemplates,
        mailings_fields: nextMailingsFields,
        updated_at: new Date().toISOString(),
      },
    };
    saveTemplatesConfig(nextConfig);
    setTemplatesConfig(nextConfig);
    setEditingTab1FieldKey(null);
    setUseStatusMessage({
      type: 'success',
      text: `Đã cập nhật trường thành "${cleanNewLabel}" (${formatFieldToken(
        cleanNewKey,
        activeSym
      )})!`,
    });
  };

  // Remove a field completely from the current template in Tab 1
  const handleRemoveFieldInTab1 = (targetKey: string) => {
    if (!chosenTemplate || !currentTemplateInfo) return;
    const activeSym = currentTemplateSymbol;

    const nextPTemplates: Record<string, string> = {};
    Object.entries(currentTemplateInfo.paragraph_templates || {}).forEach(([idxStr, tpl]) => {
      nextPTemplates[idxStr] = removeFieldKeyFromTemplateString(tpl, targetKey, activeSym);
    });

    const nextMailingsFields = { ...(currentTemplateInfo.mailings_fields || {}) };
    delete nextMailingsFields[targetKey];

    const nextFieldVals = { ...mailingsFieldValues };
    delete nextFieldVals[targetKey];
    setMailingsFieldValues(nextFieldVals);

    const nextConfig: TemplatesConfigMap = {
      ...templatesConfig,
      [chosenTemplate]: {
        ...currentTemplateInfo,
        paragraph_templates: nextPTemplates,
        mailings_fields: nextMailingsFields,
        updated_at: new Date().toISOString(),
      },
    };
    saveTemplatesConfig(nextConfig);
    setTemplatesConfig(nextConfig);
    setEditingTab1FieldKey(null);
    setUseStatusMessage({
      type: 'success',
      text: `Đã xóa trường ${formatFieldToken(targetKey, activeSym)} khỏi mẫu "${chosenTemplate}"!`,
    });
  };

  // Update field symbol in Tab 2 (Configure Template) and convert all tokens in paragraphTemplatesInput
  const handleUpdateFieldSymbolInConfig = (newSymbol: FieldSymbolPair) => {
    const prevSymbol = fieldSymbol;
    setFieldSymbol(newSymbol);
    saveFieldSymbolConfig(newSymbol);

    setParagraphTemplatesInput((prev) => {
      const next: Record<string, string> = {};
      Object.entries(prev).forEach(([idxStr, tpl]) => {
        next[idxStr] = convertTemplateTokensToSymbol(tpl, newSymbol, prevSymbol);
      });
      return next;
    });

    setConfigStatusMessage({
      type: 'success',
      text: `Đã đổi ký hiệu bao quanh trường sang ${newSymbol.openSymbol}Ten_Truong${newSymbol.closeSymbol}!`,
    });
  };

  const handleAddCustomMailingsField = (newField: MailingsFieldDef) => {
    setCustomMailingsFields((prev) => {
      const exists = prev.some((f) => f.key === newField.key);
      const next = exists
        ? prev.map((f) => (f.key === newField.key ? newField : f))
        : [...prev, newField];
      saveCustomMailingsFields(next);
      return next;
    });
  };

  const handleRenameOrUpdateMailingsField = (
    oldKey: string,
    updatedField: MailingsFieldDef
  ) => {
    setCustomMailingsFields((prev) => {
      const existsOld = prev.some((f) => f.key === oldKey);
      let next: MailingsFieldDef[];
      if (existsOld) {
        next = prev
          .filter((f) => f.key === oldKey || f.key !== updatedField.key)
          .map((f) => (f.key === oldKey ? updatedField : f));
      } else {
        const existsNew = prev.some((f) => f.key === updatedField.key);
        next = existsNew
          ? prev.map((f) => (f.key === updatedField.key ? updatedField : f))
          : [...prev, updatedField];
      }
      saveCustomMailingsFields(next);
      return next;
    });

    // Also update any occurrences of oldKey in the currently configured template (Tab 2)
    if (oldKey !== updatedField.key) {
      setParagraphTemplatesInput((prev) => {
        const next: Record<string, string> = {};
        Object.entries(prev).forEach(([idxStr, tpl]) => {
          next[idxStr] = replaceFieldKeyInTemplateString(
            tpl,
            oldKey,
            updatedField.key,
            fieldSymbol
          );
        });
        return next;
      });

      setDetectedDefaultValues((prev) => {
        const next = { ...prev };
        if (next[oldKey] !== undefined) {
          next[updatedField.key] = next[oldKey];
          delete next[oldKey];
        }
        return next;
      });
    }

    setConfigStatusMessage({
      type: 'success',
      text: `Đã cập nhật trường "${updatedField.label}" (${formatFieldToken(
        updatedField.key,
        fieldSymbol
      )})!`,
    });
  };

  const handleDeleteMailingsField = (keyToDelete: string) => {
    setCustomMailingsFields((prev) => {
      const next = prev.filter((f) => f.key !== keyToDelete);
      saveCustomMailingsFields(next);
      return next;
    });
  };

  const handleResetDefaultMailingsPalette = () => {
    const defaults = resetDefaultMailingsFields();
    setCustomMailingsFields(defaults);
    setConfigStatusMessage({
      type: 'success',
      text: 'Đã khôi phục bộ trường Mailings mặc định!',
    });
  };

  const handleSaveTemplateToLibrary = () => {
    const trimmedName = templateNameInput.trim();
    if (!trimmedName) {
      setConfigStatusMessage({
        type: 'warning',
        text: 'Vui lòng nhập tên gợi nhớ cho biểu mẫu trước khi lưu!',
      });
      return;
    }
    if (!uploadedBase64 || uploadedNonEmptyParagraphs.length === 0) {
      setConfigStatusMessage({
        type: 'warning',
        text: 'Vui lòng tải lên file Word mẫu (.docx) trước khi lưu!',
      });
      return;
    }
    if (chosenIndices.length === 0) {
      setConfigStatusMessage({
        type: 'warning',
        text: 'Vui lòng chọn ít nhất một dòng hoặc chèn ít nhất một trường dữ liệu vào mẫu!',
      });
      return;
    }

    const safeFileName = toSafeDocxFileName(trimmedName);
    const savedFilePath = `templates_store/${safeFileName}`;
    const sortedIndices = [...chosenIndices].sort((a, b) => Number(a) - Number(b));

    const mailingsFieldsMeta: NonNullable<TemplatesConfigMap[string]['mailings_fields']> = {};
    const cleanedParagraphTemplates: Record<string, string> = {};

    sortedIndices.forEach((idxStr) => {
      const tpl = paragraphTemplatesInput[idxStr];
      if (tpl) {
        cleanedParagraphTemplates[idxStr] = tpl;
        const keys = extractMergeFieldKeysFromText(tpl, fieldSymbol);
        keys.forEach((k) => {
          const def = getFieldDefByKey(k, customMailingsFields);
          mailingsFieldsMeta[k] = {
            label: def.label,
            defaultValue: detectedDefaultValues[k] ?? def.placeholder ?? '',
            category: def.category,
          };
        });
      }
    });

    const updatedConfig: TemplatesConfigMap = { ...templatesConfig };
    if (editingExistingTemplateName && editingExistingTemplateName !== trimmedName) {
      delete updatedConfig[editingExistingTemplateName];
    }

    updatedConfig[trimmedName] = {
      file_path: savedFilePath,
      selected_indices: sortedIndices,
      paragraph_templates: cleanedParagraphTemplates,
      mailings_fields: mailingsFieldsMeta,
      field_symbol: fieldSymbol,
      custom_labels: customLabelsInput,
      base64_docx: uploadedBase64,
      updated_at: new Date().toISOString(),
      total_paragraphs: uploadedAllParagraphs.length,
      category: templateCategoryInput.trim() || 'Biểu mẫu chung',
    };

    saveTemplatesConfig(updatedConfig);
    setTemplatesConfig(updatedConfig);
    setChosenTemplate(trimmedName);
    setEditingExistingTemplateName(trimmedName);

    const totalFieldsCount = Object.keys(mailingsFieldsMeta).length;
    setConfigStatusMessage({
      type: 'success',
      text: `Đã lưu thành công mẫu "${trimmedName}" (${totalFieldsCount} trường ${fieldSymbol.openSymbol}...${fieldSymbol.closeSymbol} trên ${sortedIndices.length} dòng) vào thư viện!`,
    });
  };

  // Export & download filled Word file locally in Tab 1
  const handleExportFilledWord = async () => {
    if (!chosenTemplate || !currentTemplateInfo) return;
    setIsExporting(true);
    setUseStatusMessage(null);

    try {
      const blob = await exportFilledDocxBlob(currentTemplateInfo.base64_docx, userReplacements);
      const outputFileName = `${chosenTemplate}_da_dien.docx`;

      setLastExportedBlob({
        blob,
        fileName: outputFileName,
        templateName: chosenTemplate,
      });

      triggerBlobDownload(blob, outputFileName);

      const newHistory = addExportHistory({
        id: `${Date.now()}`,
        templateName: chosenTemplate,
        outputFileName,
        exportedAt: new Date().toISOString(),
        modifiedCount: modifiedFieldsCount,
        replacements: { ...userReplacements },
        fieldValues: { ...mailingsFieldValues },
      });
      setExportHistory(newHistory);

      setUseStatusMessage({
        type: 'success',
        text: `Đã trộn thư và tải xuống thành công file Word "${outputFileName}"!`,
      });
    } catch (err) {
      setUseStatusMessage({
        type: 'error',
        text: err instanceof Error ? err.message : 'Có lỗi xảy ra khi xuất file Word.',
      });
    } finally {
      setIsExporting(false);
    }
  };

  // Save filled Word (.docx) directly to Google Drive
  const handleSaveFilledWordToDrive = async () => {
    if (!chosenTemplate || !currentTemplateInfo) return;
    const token = await getAccessToken();
    if (!token || needsAuth) {
      setIsDriveModalOpen(true);
      return;
    }

    setIsSavingToDrive(true);
    setUseStatusMessage(null);
    setLastDriveSavedFile(null);

    try {
      const blob = await exportFilledDocxBlob(currentTemplateInfo.base64_docx, userReplacements);
      const outputFileName = `${chosenTemplate}_da_dien.docx`;

      const uploaded = await uploadOrUpdateFileToDrive({
        fileName: outputFileName,
        blob,
        mimeType:
          'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      });

      setLastDriveSavedFile(uploaded);

      const newHistory = addExportHistory({
        id: `${Date.now()}`,
        templateName: chosenTemplate,
        outputFileName,
        exportedAt: new Date().toISOString(),
        modifiedCount: modifiedFieldsCount,
        replacements: { ...userReplacements },
        fieldValues: { ...mailingsFieldValues },
      });
      setExportHistory(newHistory);

      setUseStatusMessage({
        type: 'success',
        text: `Đã lưu trực tiếp file "${uploaded.name}" lên Google Drive của bạn!`,
      });
    } catch (err: any) {
      if (err?.message === 'UNAUTHENTICATED') {
        setNeedsAuth(true);
        setIsDriveModalOpen(true);
      } else {
        setUseStatusMessage({
          type: 'error',
          text:
            err instanceof Error
              ? err.message
              : 'Không thể lưu file lên Google Drive.',
        });
      }
    } finally {
      setIsSavingToDrive(false);
    }
  };

  // Upload any template or history record directly to Google Drive from Tab 3
  const handleQuickUploadDocxToDrive = async (
    templateName: string,
    replacementsMap: Record<number, string>,
    targetFileName: string
  ) => {
    const token = await getAccessToken();
    if (!token || needsAuth) {
      setIsDriveModalOpen(true);
      return;
    }

    const info = templatesConfig[templateName];
    if (!info) return;

    try {
      setLibraryBannerMessage(`Đang tải "${targetFileName}" lên Google Drive...`);
      const blob = await exportFilledDocxBlob(info.base64_docx, replacementsMap);
      const uploaded = await uploadOrUpdateFileToDrive({
        fileName: targetFileName,
        blob,
        mimeType:
          'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      });
      setLibraryBannerMessage(
        `Đã tải thành công "${uploaded.name}" lên Google Drive!`
      );
    } catch (err: any) {
      if (err?.message === 'UNAUTHENTICATED') {
        setNeedsAuth(true);
        setIsDriveModalOpen(true);
      } else {
        setLibraryBannerMessage(
          err instanceof Error ? err.message : 'Lỗi khi tải tệp lên Google Drive.'
        );
      }
    }
  };

  const handleDownloadOriginalTemplate = async (templateName: string) => {
    const info = templatesConfig[templateName];
    if (!info) return;
    const blob = await exportFilledDocxBlob(info.base64_docx, {});
    const rawFileName = info.file_path.split('/').pop() || `${templateName}_goc.docx`;
    triggerBlobDownload(blob, rawFileName);
  };

  const handleDeleteTemplate = (templateName: string) => {
    const next = { ...templatesConfig };
    delete next[templateName];
    saveTemplatesConfig(next);
    setTemplatesConfig(next);
    const remainingNames = Object.keys(next);
    if (chosenTemplate === templateName) {
      setChosenTemplate(remainingNames[0] || '');
    }
  };

  const handleResetAllFields = () => {
    if (!currentTemplateInfo) return;
    const resetFieldVals: Record<string, string> = {};
    if (currentTemplateInfo.mailings_fields) {
      Object.entries(currentTemplateInfo.mailings_fields).forEach(([k, meta]) => {
        resetFieldVals[k] = meta.defaultValue || '';
      });
    }
    setMailingsFieldValues(resetFieldVals);

    const resetMap: Record<number, string> = {};
    currentTemplateInfo.selected_indices.forEach((idxStr) => {
      const idx = Number(idxStr);
      const original = parsedTemplateParagraphs[idx]?.text ?? '';
      const tpl = currentTemplateInfo.paragraph_templates?.[idxStr];
      resetMap[idx] = tpl ? resolveParagraphTemplate(tpl, resetFieldVals) : original;
    });
    setUserReplacements(resetMap);
    setLastExportedBlob(null);
    setUseStatusMessage({
      type: 'success',
      text: 'Đã khôi phục toàn bộ các trường dữ liệu về giá trị mặc định của biểu mẫu.',
    });
  };

  const handleClearAllMailingsInputs = () => {
    if (!currentTemplateInfo) return;
    const emptyVals: Record<string, string> = {};
    activeTemplateMailingsFields.forEach((f) => {
      emptyVals[f.key] = '';
    });
    setMailingsFieldValues(emptyVals);

    const nextReplacements: Record<number, string> = { ...userReplacements };
    currentTemplateInfo.selected_indices.forEach((idxStr) => {
      const idx = Number(idxStr);
      const tpl = currentTemplateInfo.paragraph_templates?.[idxStr];
      if (tpl) {
        nextReplacements[idx] = resolveParagraphTemplate(tpl, emptyVals);
      }
    });
    setUserReplacements(nextReplacements);
  };

  const filteredConfigureParagraphs = useMemo(() => {
    const q = paragraphSearchQuery.trim().toLowerCase();
    return uploadedNonEmptyParagraphs.filter((item) => {
      if (configureSectionFilter !== 'all') {
        if (configureSectionFilter === 'footnote') {
          if (item.section !== 'footnote' && item.section !== 'endnote') return false;
        } else if (item.section !== configureSectionFilter) {
          return false;
        }
      }
      if (!q) return true;
      const tpl = paragraphTemplatesInput[String(item.index)] || item.text;
      return (
        item.text.toLowerCase().includes(q) ||
        tpl.toLowerCase().includes(q) ||
        item.sectionLabel.toLowerCase().includes(q) ||
        (item.tablePosition && item.tablePosition.toLowerCase().includes(q)) ||
        `dòng ${item.index}`.includes(q) ||
        String(item.index) === q
      );
    });
  }, [uploadedNonEmptyParagraphs, paragraphTemplatesInput, paragraphSearchQuery, configureSectionFilter]);

  const activeConfigParagraph = useMemo(() => {
    if (activeConfigParagraphIndex === null) return null;
    return (
      uploadedAllParagraphs.find((p) => p.index === activeConfigParagraphIndex) || null
    );
  }, [uploadedAllParagraphs, activeConfigParagraphIndex]);

  const activeConfigTemplateString = useMemo(() => {
    if (!activeConfigParagraph) return '';
    return (
      paragraphTemplatesInput[String(activeConfigParagraph.index)] ??
      activeConfigParagraph.text
    );
  }, [activeConfigParagraph, paragraphTemplatesInput]);

  return (
    <div className="min-h-screen flex flex-col bg-[#F8FAFC] text-slate-900">
      {/* Top Bar Contract: Strictly 3 zones separated by gap-8 */}
      <header className="sticky top-0 z-30 flex items-center justify-between gap-8 px-6 lg:px-10 py-3.5 bg-white border-b border-slate-200">
        {/* Zone 1: Single text element wordmark */}
        <a
          href="#top"
          onClick={(e) => {
            e.preventDefault();
            setActiveSection('library');
          }}
          className="text-lg font-bold tracking-tight text-slate-900 font-display whitespace-nowrap shrink-0"
        >
          Thư Viện Biểu Mẫu Word
        </a>

        {/* Zone 2: 4 clean single-line text navigation links */}
        <nav className="hidden md:flex items-center gap-7 text-sm font-medium text-slate-600">
          <button
            type="button"
            onClick={() => setActiveSection('library')}
            className={`py-1 transition-colors whitespace-nowrap shrink-0 border-b-2 cursor-pointer ${
              activeSection === 'library'
                ? 'text-blue-700 border-blue-700 font-semibold'
                : 'border-transparent hover:text-slate-900'
            }`}
          >
            Kho biểu mẫu ({templateNames.length})
          </button>
          <button
            type="button"
            onClick={() => setActiveSection('use')}
            className={`py-1 transition-colors whitespace-nowrap shrink-0 border-b-2 cursor-pointer ${
              activeSection === 'use'
                ? 'text-blue-700 border-blue-700 font-semibold'
                : 'border-transparent hover:text-slate-900'
            }`}
          >
            Sử dụng mẫu (Mailings)
          </button>
          <button
            type="button"
            onClick={() => setActiveSection('configure')}
            className={`py-1 transition-colors whitespace-nowrap shrink-0 border-b-2 cursor-pointer ${
              activeSection === 'configure'
                ? 'text-blue-700 border-blue-700 font-semibold'
                : 'border-transparent hover:text-slate-900'
            }`}
          >
            Thêm & Cấu hình vị trí trường
          </button>
          <button
            type="button"
            onClick={() => setActiveSection('json')}
            className={`py-1 transition-colors whitespace-nowrap shrink-0 border-b-2 cursor-pointer ${
              activeSection === 'json'
                ? 'text-blue-700 border-blue-700 font-semibold'
                : 'border-transparent hover:text-slate-900'
            }`}
          >
            Cấu hình JSON
          </button>
        </nav>

        {/* Zone 3: 1 primary action */}
        <div className="flex items-center shrink-0">
          {needsAuth || !googleUser ? (
            <GoogleSignInButton
              onClick={handleGoogleLogin}
              disabled={isLoggingIn}
              label={isLoggingIn ? 'Signing in...' : 'Sign in with Google'}
            />
          ) : (
            <button
              type="button"
              onClick={() => setIsDriveModalOpen(true)}
              className="inline-flex items-center gap-2 px-4 py-2 text-xs font-semibold text-white bg-blue-700 rounded-lg hover:bg-blue-800 transition-colors whitespace-nowrap shrink-0 cursor-pointer"
            >
              <Cloud className="w-3.5 h-3.5" />
              <span>Google Drive ({googleUser.email?.split('@')[0]})</span>
            </button>
          )}
        </div>
      </header>

      {/* Google Drive Modal */}
      <GoogleDriveModal
        isOpen={isDriveModalOpen}
        onClose={() => setIsDriveModalOpen(false)}
        user={googleUser}
        needsAuth={needsAuth}
        isLoggingIn={isLoggingIn}
        onLogin={handleGoogleLogin}
        onLogout={handleGoogleLogout}
        templatesConfig={templatesConfig}
        sharedDriveInfo={sharedDriveInfo}
        onUpdateSharedDriveInfo={(info, updatedTemplates) => {
          setSharedDriveInfo(info);
          setTemplatesConfig(updatedTemplates);
          saveTemplatesConfig(updatedTemplates);
          if (!chosenTemplate && Object.keys(updatedTemplates).length > 0) {
            setChosenTemplate(Object.keys(updatedTemplates)[0]);
          }
        }}
        onImportDocxFromDrive={async (file) => {
          setEditingExistingTemplateName(null);
          setActiveSection('configure');
          await handleFileUpload(file);
        }}
        onRestoreLibraryFromDrive={(restored) => {
          const merged = { ...templatesConfig, ...restored };
          saveTemplatesConfig(merged);
          setTemplatesConfig(merged);
          if (!chosenTemplate && Object.keys(merged).length > 0) {
            setChosenTemplate(Object.keys(merged)[0]);
          }
        }}
        onAuthExpired={() => {
          setNeedsAuth(true);
        }}
      />

      {/* Mandatory Confirmation Dialog for Publishing & Sharing Library on Google Drive */}
      {isConfirmingDrivePublish && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/50 px-4">
          <div className="w-full max-w-lg bg-white border border-slate-200 rounded-lg p-6 space-y-4">
            <div className="flex items-start gap-3">
              <AlertTriangle className="w-5 h-5 text-amber-600 shrink-0 mt-0.5" />
              <div className="space-y-1">
                <h3 className="text-base font-bold text-slate-900">
                  Xác nhận Lưu & Chia Sẻ Kho Biểu Mẫu trên Google Drive?
                </h3>
                <p className="text-xs text-slate-600 leading-relaxed">
                  Ứng dụng sẽ tạo/cập nhật thư mục{' '}
                  <strong className="text-slate-900">&ldquo;{SHARED_DRIVE_FOLDER_NAME}&rdquo;</strong> trên Google Drive của bạn ({googleUser?.email}), tải lên toàn bộ{' '}
                  <strong className="text-slate-900">{templateNames.length} file Word (.docx)</strong> cùng tệp dữ liệu{' '}
                  <code className="font-mono-tabular text-blue-700">{SHARED_BACKUP_FILE_NAME}</code>, và bật quyền{' '}
                  <strong className="text-slate-900">&ldquo;Bất kỳ ai có đường liên kết đều có thể xem&rdquo; (Anyone with the link can view)</strong>.
                </p>
                <p className="text-xs text-emerald-800 bg-emerald-50 border border-emerald-200 rounded p-2.5 mt-2">
                  Sau khi xác nhận, bất kỳ người dùng nào truy cập vào ứng dụng (kể cả không đăng nhập tài khoản Google của bạn) đều sẽ tự động nhìn thấy và sử dụng được toàn bộ các mẫu lưu trên Google Drive này.
                </p>
              </div>
            </div>

            <div className="flex items-center justify-end gap-2.5 pt-2 border-t border-slate-100">
              <button
                type="button"
                onClick={() => setIsConfirmingDrivePublish(false)}
                className="px-3.5 py-2 text-xs font-medium text-slate-700 bg-slate-100 hover:bg-slate-200 rounded-lg cursor-pointer"
              >
                Hủy bỏ
              </button>
              <button
                type="button"
                onClick={handleConfirmPublishToSharedDrive}
                className="inline-flex items-center gap-1.5 px-4 py-2 text-xs font-semibold text-white bg-blue-700 hover:bg-blue-800 rounded-lg cursor-pointer"
              >
                <CloudUpload className="w-3.5 h-3.5" />
                <span>Xác nhận Lưu & Chia Sẻ trên Drive</span>
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Duplicate Template Modal */}
      {duplicatingSourceName && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/45 px-4">
          <div className="w-full max-w-md bg-white border border-slate-200 rounded-lg p-6 space-y-4">
            <div className="flex items-start justify-between gap-3">
              <div>
                <h3 className="text-base font-semibold text-slate-900 flex items-center gap-2">
                  <CopyPlus className="w-4 h-4 text-blue-700" />
                  <span>Nhân bản biểu mẫu Word</span>
                </h3>
                <p className="text-xs text-slate-500 mt-0.5">
                  Sao chép toàn bộ file Word gốc và vị trí các trường «Mailings» từ mẫu &ldquo;{duplicatingSourceName}&rdquo;
                </p>
              </div>
              <button
                type="button"
                onClick={() => setDuplicatingSourceName(null)}
                className="p-1 text-slate-400 hover:text-slate-700 rounded cursor-pointer"
              >
                <X className="w-4 h-4" />
              </button>
            </div>

            <div className="space-y-3">
              <div>
                <label className="block text-xs font-medium text-slate-700 mb-1">
                  Tên biểu mẫu bản sao mới:
                </label>
                <input
                  type="text"
                  value={duplicateNewName}
                  onChange={(e) => {
                    setDuplicateNewName(e.target.value);
                    setDuplicateError(null);
                  }}
                  placeholder="Nhập tên cho mẫu nhân bản..."
                  className="w-full px-3 py-2 text-sm bg-slate-50 border border-slate-300 rounded-lg text-slate-900 focus:bg-white focus:outline-none focus:border-blue-600"
                  autoFocus
                />
              </div>

              <div>
                <label className="block text-xs font-medium text-slate-700 mb-1">
                  Phân nhóm biểu mẫu:
                </label>
                <input
                  type="text"
                  value={duplicateNewCategory}
                  onChange={(e) => setDuplicateNewCategory(e.target.value)}
                  placeholder="VD: Hành chính - Nhân sự, Kế toán..."
                  className="w-full px-3 py-2 text-sm bg-slate-50 border border-slate-300 rounded-lg text-slate-900 focus:bg-white focus:outline-none focus:border-blue-600"
                />
              </div>

              <label className="flex items-center gap-2 text-xs text-slate-700 cursor-pointer select-none pt-1">
                <input
                  type="checkbox"
                  checked={openConfigAfterDuplicate}
                  onChange={(e) => setOpenConfigAfterDuplicate(e.target.checked)}
                  className="rounded border-slate-300 text-blue-700 focus:ring-blue-600"
                />
                <span>Mở màn hình chỉnh sửa vị trí «Trường» ngay sau khi nhân bản</span>
              </label>

              {duplicateError && (
                <p className="text-xs font-medium text-red-700 bg-red-50 border border-red-200 rounded p-2.5">
                  {duplicateError}
                </p>
              )}
            </div>

            <div className="flex items-center justify-end gap-2.5 pt-2 border-t border-slate-100">
              <button
                type="button"
                onClick={() => setDuplicatingSourceName(null)}
                className="px-3.5 py-2 text-xs font-medium text-slate-700 bg-slate-100 hover:bg-slate-200 rounded-lg whitespace-nowrap shrink-0 cursor-pointer"
              >
                Hủy bỏ
              </button>
              <button
                type="button"
                onClick={handleConfirmDuplicateTemplate}
                className="inline-flex items-center gap-1.5 px-4 py-2 text-xs font-semibold text-white bg-blue-700 hover:bg-blue-800 rounded-lg whitespace-nowrap shrink-0 cursor-pointer"
              >
                <CopyPlus className="w-3.5 h-3.5" />
                <span>Tạo bản sao</span>
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Main Content Container */}
      <main className="flex-1 w-full max-w-[1440px] mx-auto px-4 sm:px-6 lg:px-10 py-6 lg:py-8">
        {/* Workspace Header & Segmented Tab Switcher */}
        <div className="flex flex-col md:flex-row md:items-end justify-between gap-4 pb-6 mb-6 border-b border-slate-200">
          <div>
            <h1 className="text-2xl sm:text-3xl font-bold tracking-tight text-slate-900 font-display">
              Thư Viện Biểu Mẫu & Trộn Thư Mailings Word
            </h1>
            <div className="flex flex-wrap items-center gap-2 mt-1.5 text-xs text-slate-500">
              <span>Chèn trường dữ liệu mặc định «Trường» vào đúng vị trí trong câu Word</span>
              <span aria-hidden="true">·</span>
              <span className="font-mono-tabular">{templateNames.length} mẫu đã lưu</span>
              <span aria-hidden="true">·</span>
              <span className="font-mono-tabular">{exportHistory.length} lượt xuất file</span>
              <span aria-hidden="true">·</span>
              <button
                type="button"
                onClick={() => setIsDriveModalOpen(true)}
                className="text-blue-700 hover:underline font-medium cursor-pointer"
              >
                {googleUser && !needsAuth
                  ? `Đã kết nối Google Drive (${googleUser.email})`
                  : 'Kết nối Google Drive'}
              </button>
            </div>
          </div>

          {/* Interactive Segmented Control Tabs */}
          <div className="flex items-center gap-1 p-1 bg-slate-200/75 rounded-lg self-start md:self-auto overflow-x-auto max-w-full">
            <button
              type="button"
              onClick={() => setActiveSection('library')}
              className={`flex items-center gap-1.5 px-3.5 py-2 text-xs font-semibold rounded-md transition-colors whitespace-nowrap shrink-0 cursor-pointer ${
                activeSection === 'library'
                  ? 'bg-white text-slate-900 shadow-xs'
                  : 'text-slate-600 hover:text-slate-900'
              }`}
            >
              <FolderOpen className="w-3.5 h-3.5 text-blue-700" />
              <span>Kho Mẫu & Lịch Sử</span>
            </button>
            <button
              type="button"
              onClick={() => setActiveSection('use')}
              className={`flex items-center gap-1.5 px-3.5 py-2 text-xs font-semibold rounded-md transition-colors whitespace-nowrap shrink-0 cursor-pointer ${
                activeSection === 'use'
                  ? 'bg-white text-slate-900 shadow-xs'
                  : 'text-slate-600 hover:text-slate-900'
              }`}
            >
              <FileText className="w-3.5 h-3.5 text-blue-700" />
              <span>Sử Dụng Mẫu Đã Lưu</span>
            </button>
            <button
              type="button"
              onClick={() => setActiveSection('configure')}
              className={`flex items-center gap-1.5 px-3.5 py-2 text-xs font-semibold rounded-md transition-colors whitespace-nowrap shrink-0 cursor-pointer ${
                activeSection === 'configure'
                  ? 'bg-white text-slate-900 shadow-xs'
                  : 'text-slate-600 hover:text-slate-900'
              }`}
            >
              <Plus className="w-3.5 h-3.5 text-blue-700" />
              <span>Thêm / Cấu Hình Vị Trí Trường Mới</span>
            </button>
            <button
              type="button"
              onClick={() => setActiveSection('json')}
              className={`flex items-center gap-1.5 px-3.5 py-2 text-xs font-semibold rounded-md transition-colors whitespace-nowrap shrink-0 cursor-pointer ${
                activeSection === 'json'
                  ? 'bg-white text-slate-900 shadow-xs'
                  : 'text-slate-600 hover:text-slate-900'
              }`}
            >
              <SlidersHorizontal className="w-3.5 h-3.5 text-slate-700" />
              <span>config.json</span>
            </button>
          </div>
        </div>

        {/* ================= TAB 1: SỬ DỤNG MẪU ĐÃ LƯU ================= */}
        {activeSection === 'use' && (
          <div>
            {isInitializing ? (
              <div className="p-12 bg-white border border-slate-200 rounded-lg text-center text-sm text-slate-500">
                Đang tải thư viện biểu mẫu Word...
              </div>
            ) : templateNames.length === 0 ? (
              <div className="p-10 bg-white border border-slate-200 rounded-lg text-center max-w-xl mx-auto my-8">
                <h2 className="text-lg font-semibold text-slate-900">
                  Chưa có mẫu nào được lưu trong thư viện
                </h2>
                <p className="mt-2 text-sm text-slate-600">
                  Hãy chuyển sang tab &ldquo;Thêm / Cấu Hình Vị Trí Trường Mới&rdquo; hoặc tải biểu mẫu từ Google Drive.
                </p>
                <div className="mt-5 flex items-center justify-center gap-3">
                  <button
                    type="button"
                    onClick={() => setActiveSection('configure')}
                    className="px-4 py-2 text-xs font-semibold text-white bg-blue-700 rounded-lg hover:bg-blue-800 transition-colors whitespace-nowrap shrink-0 cursor-pointer"
                  >
                    Thêm mẫu đầu tiên
                  </button>
                  <button
                    type="button"
                    onClick={() => setIsDriveModalOpen(true)}
                    className="inline-flex items-center gap-1.5 px-4 py-2 text-xs font-semibold text-slate-700 bg-slate-100 hover:bg-slate-200 rounded-lg transition-colors whitespace-nowrap shrink-0 cursor-pointer"
                  >
                    <Cloud className="w-3.5 h-3.5 text-blue-700" />
                    <span>Chọn từ Google Drive</span>
                  </button>
                </div>
              </div>
            ) : (
              <div className="grid grid-cols-1 lg:grid-cols-12 gap-8 items-start">
                {/* Left Column (5 cols on lg): Template Selector & Mailings Field Inputs */}
                <div className="lg:col-span-5 space-y-6">
                  {/* Template Selector Box */}
                  <section className="p-5 bg-white border border-slate-200 rounded-lg space-y-4">
                    <div className="flex flex-wrap items-center justify-between gap-2">
                      <h2 className="text-base font-semibold text-slate-900">
                        Chọn biểu mẫu để trộn thư & điền nhanh
                      </h2>
                      {chosenTemplate && (
                        <div className="flex flex-wrap items-center gap-3">
                          <button
                            type="button"
                            onClick={() => handleShareSingleTemplateLink(chosenTemplate)}
                            className="inline-flex items-center gap-1 text-xs font-semibold text-emerald-700 hover:text-emerald-800 whitespace-nowrap shrink-0 cursor-pointer"
                          >
                            <Share2 className="w-3.5 h-3.5" />
                            <span>Chia sẻ link mẫu</span>
                          </button>
                          <button
                            type="button"
                            onClick={() => handleOpenDuplicateModal(chosenTemplate)}
                            className="inline-flex items-center gap-1 text-xs font-semibold text-slate-700 hover:text-blue-700 whitespace-nowrap shrink-0 cursor-pointer"
                          >
                            <CopyPlus className="w-3.5 h-3.5" />
                            <span>Nhân bản</span>
                          </button>
                          <button
                            type="button"
                            onClick={() => handleEditExistingTemplateConfig(chosenTemplate)}
                            className="inline-flex items-center gap-1 text-xs font-semibold text-blue-700 hover:text-blue-800 whitespace-nowrap shrink-0 cursor-pointer"
                          >
                            <Settings2 className="w-3.5 h-3.5" />
                            <span>Sửa vị trí «Trường»</span>
                          </button>
                        </div>
                      )}
                    </div>

                    <div>
                      <label
                        htmlFor="select-template"
                        className="block text-xs font-medium text-slate-700 mb-1.5"
                      >
                        Chọn mẫu Word:
                      </label>
                      <select
                        id="select-template"
                        value={chosenTemplate}
                        onChange={(e) => setChosenTemplate(e.target.value)}
                        className="w-full px-3.5 py-2.5 text-sm bg-slate-50 border border-slate-300 rounded-lg text-slate-900 focus:outline-none focus:border-blue-600 focus:bg-white transition-colors"
                      >
                        {templateNames.map((name) => {
                          const entry = templatesConfig[name];
                          const fieldCount = entry.mailings_fields
                            ? Object.keys(entry.mailings_fields).length
                            : 0;
                          return (
                            <option key={name} value={name}>
                              {name} ({fieldCount > 0 ? `${fieldCount} trường Mailings · ` : ''}
                              {entry.selected_indices.length} dòng)
                            </option>
                          );
                        })}
                      </select>
                    </div>

                    {currentTemplateInfo && (
                      <div className="pt-3 border-t border-slate-100 flex flex-wrap items-center justify-between gap-2 text-xs text-slate-600">
                        <div className="flex flex-wrap items-center gap-1.5">
                          <span className="font-medium text-slate-800">
                            {activeTemplateMailingsFields.length} trường Mailings
                          </span>
                          <span aria-hidden="true">·</span>
                          <span className="font-mono-tabular">
                            {currentTemplateInfo.selected_indices.length} dòng văn bản
                          </span>
                        </div>
                        <button
                          type="button"
                          onClick={() => handleDownloadOriginalTemplate(chosenTemplate)}
                          className="inline-flex items-center gap-1 text-xs font-medium text-slate-600 hover:text-slate-900 whitespace-nowrap shrink-0 cursor-pointer"
                        >
                          <Download className="w-3.5 h-3.5" />
                          <span>Tải file mẫu gốc</span>
                        </button>
                      </div>
                    )}
                  </section>

                  {/* Dynamic Inputs Box */}
                  {currentTemplateInfo && (
                    <section className="p-5 bg-white border border-slate-200 rounded-lg space-y-5">
                      <div className="flex flex-col gap-3 pb-3 border-b border-slate-200">
                        <div className="flex flex-wrap items-center justify-between gap-2">
                          <div className="flex flex-wrap items-center gap-1 p-1 bg-slate-100 rounded-lg">
                            <button
                              type="button"
                              onClick={() => setFillMode('mailings_fields')}
                              className={`px-2.5 py-1.5 text-xs font-semibold rounded-md transition-colors whitespace-nowrap shrink-0 cursor-pointer ${
                                fillMode === 'mailings_fields'
                                  ? 'bg-white text-blue-800 shadow-2xs'
                                  : 'text-slate-600 hover:text-slate-900'
                              }`}
                            >
                              Điền «Trường Mailings» ({activeTemplateMailingsFields.length})
                            </button>
                            <button
                              type="button"
                              onClick={() => setFillMode('full_lines')}
                              className={`px-2.5 py-1.5 text-xs font-semibold rounded-md transition-colors whitespace-nowrap shrink-0 cursor-pointer ${
                                fillMode === 'full_lines'
                                  ? 'bg-white text-blue-800 shadow-2xs'
                                  : 'text-slate-600 hover:text-slate-900'
                              }`}
                            >
                              Dòng đã chọn ({currentTemplateInfo.selected_indices.length})
                            </button>
                            <button
                              type="button"
                              onClick={() => setFillMode('all_document_lines')}
                              className={`px-2.5 py-1.5 text-xs font-semibold rounded-md transition-colors whitespace-nowrap shrink-0 cursor-pointer ${
                                fillMode === 'all_document_lines'
                                  ? 'bg-white text-blue-800 shadow-2xs'
                                  : 'text-slate-600 hover:text-slate-900'
                              }`}
                            >
                              Toàn bộ văn bản ({parsedTemplateParagraphs.filter((p) => !p.isEmpty).length})
                            </button>
                          </div>

                          <div className="flex items-center gap-2">
                            {fillMode === 'mailings_fields' && activeTemplateMailingsFields.length > 0 && (
                              <button
                                type="button"
                                onClick={handleClearAllMailingsInputs}
                                className="px-2.5 py-1 text-xs font-medium text-slate-600 bg-slate-100 hover:bg-slate-200 rounded-md transition-colors whitespace-nowrap shrink-0 cursor-pointer"
                              >
                                Xóa trắng ô
                              </button>
                            )}
                            <button
                              type="button"
                              onClick={handleResetAllFields}
                              className="inline-flex items-center gap-1 px-2.5 py-1 text-xs font-medium text-slate-600 bg-slate-100 hover:bg-slate-200 rounded-md transition-colors whitespace-nowrap shrink-0 cursor-pointer"
                            >
                              <RotateCcw className="w-3 h-3" />
                              <span>Mặc định</span>
                            </button>
                          </div>
                        </div>

                        {fillMode === 'all_document_lines' && (
                          <div className="flex flex-wrap items-center gap-1 pt-1">
                            {[
                              { id: 'all', label: 'Tất cả vùng' },
                              { id: 'header', label: 'Headnote (Đầu trang)' },
                              { id: 'body', label: 'Nội dung & Bảng' },
                              { id: 'footnote', label: 'Footnote (Chú thích)' },
                              { id: 'footer', label: 'Footer (Chân trang)' },
                            ].map((sec) => (
                              <button
                                key={sec.id}
                                type="button"
                                onClick={() => setUseSectionFilter(sec.id as any)}
                                className={`px-2 py-1 text-[11px] font-medium rounded transition-colors cursor-pointer ${
                                  useSectionFilter === sec.id
                                    ? 'bg-slate-900 text-white'
                                    : 'bg-slate-100 text-slate-600 hover:text-slate-900'
                                }`}
                              >
                                {sec.label}
                              </button>
                            ))}
                          </div>
                        )}
                      </div>

                      {fillMode === 'mailings_fields' ? (
                        activeTemplateMailingsFields.length === 0 ? (
                          <div className="p-6 text-center bg-slate-50 rounded-lg border border-slate-200 space-y-2">
                            <p className="text-xs font-medium text-slate-700">
                              Mẫu này chưa gắn trường «Mailings» vào vị trí cụ thể trong câu.
                            </p>
                            <div className="flex items-center justify-center gap-2">
                              <button
                                type="button"
                                onClick={() => handleEditExistingTemplateConfig(chosenTemplate)}
                                className="px-3 py-1.5 text-xs font-semibold text-white bg-blue-700 hover:bg-blue-800 rounded-md cursor-pointer"
                              >
                                Gắn vị trí trường Mailings ngay
                              </button>
                              <button
                                type="button"
                                onClick={() => setFillMode('all_document_lines')}
                                className="px-3 py-1.5 text-xs font-medium text-slate-700 bg-white border border-slate-300 rounded-md cursor-pointer"
                              >
                                Sửa toàn bộ văn bản (Headnote, Body, Footnote)
                              </button>
                            </div>
                          </div>
                        ) : (
                          <div className="space-y-3">
                            {/* Quick Symbol Delimiter Bar in Tab 1 */}
                            <div className="flex flex-wrap items-center justify-between gap-2 px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg text-xs">
                              <div className="flex items-center gap-1.5 text-slate-600">
                                <Braces className="w-3.5 h-3.5 text-blue-700" />
                                <span>
                                  Ký hiệu trường đang dùng:{' '}
                                  <strong className="font-mono-tabular text-blue-800">
                                    {formatFieldToken('Ten_Truong', currentTemplateSymbol)}
                                  </strong>
                                </span>
                              </div>
                              <button
                                type="button"
                                onClick={() => {
                                  setTab1CustomOpen(currentTemplateSymbol.openSymbol);
                                  setTab1CustomClose(currentTemplateSymbol.closeSymbol);
                                  setShowTab1SymbolPicker((prev) => !prev);
                                }}
                                className="text-xs font-semibold text-blue-700 hover:underline cursor-pointer"
                              >
                                {showTab1SymbolPicker ? 'Đóng tùy chỉnh ký hiệu' : 'Đổi ký hiệu trường'}
                              </button>
                            </div>

                            {showTab1SymbolPicker && (
                              <div className="p-3 bg-blue-50/40 border border-blue-200 rounded-lg space-y-2.5">
                                <div className="flex flex-wrap items-center gap-1.5">
                                  {PRESET_FIELD_SYMBOLS.map((preset) => {
                                    const isSel =
                                      currentTemplateSymbol.openSymbol === preset.openSymbol &&
                                      currentTemplateSymbol.closeSymbol === preset.closeSymbol;
                                    return (
                                      <button
                                        key={preset.id}
                                        type="button"
                                        onClick={() =>
                                          handleUpdateFieldSymbolInTab1({
                                            openSymbol: preset.openSymbol,
                                            closeSymbol: preset.closeSymbol,
                                          })
                                        }
                                        className={`px-2 py-1 text-[11px] rounded border transition-colors cursor-pointer ${
                                          isSel
                                            ? 'bg-blue-700 text-white border-blue-700 font-semibold'
                                            : 'bg-white text-slate-700 border-slate-200 hover:border-blue-400'
                                        }`}
                                      >
                                        {preset.name}
                                      </button>
                                    );
                                  })}
                                </div>
                                <div className="flex flex-wrap items-center gap-2 pt-1">
                                  <span className="text-[11px] text-slate-600">Tự nhập ký hiệu:</span>
                                  <input
                                    type="text"
                                    value={tab1CustomOpen}
                                    onChange={(e) => setTab1CustomOpen(e.target.value)}
                                    placeholder="Mở (VD: «)"
                                    className="w-16 px-2 py-1 text-xs font-mono-tabular bg-white border border-slate-300 rounded"
                                  />
                                  <input
                                    type="text"
                                    value={tab1CustomClose}
                                    onChange={(e) => setTab1CustomClose(e.target.value)}
                                    placeholder="Đóng (VD: »)"
                                    className="w-16 px-2 py-1 text-xs font-mono-tabular bg-white border border-slate-300 rounded"
                                  />
                                  <button
                                    type="button"
                                    onClick={() => {
                                      if (tab1CustomOpen.trim() && tab1CustomClose.trim()) {
                                        handleUpdateFieldSymbolInTab1({
                                          openSymbol: tab1CustomOpen.trim(),
                                          closeSymbol: tab1CustomClose.trim(),
                                        });
                                      }
                                    }}
                                    className="px-2.5 py-1 text-xs font-semibold text-white bg-blue-700 hover:bg-blue-800 rounded cursor-pointer"
                                  >
                                    Áp dụng
                                  </button>
                                </div>
                              </div>
                            )}

                            <div className="space-y-3.5 max-h-[520px] overflow-y-auto pr-1">
                              {activeTemplateMailingsFields.map((field) => {
                                const val = mailingsFieldValues[field.key] ?? '';
                                const isFocused = activeFieldKey === field.key;
                                const isEditingThisField = editingTab1FieldKey === field.key;
                                const formattedToken = formatFieldToken(field.key, currentTemplateSymbol);

                                return (
                                  <div
                                    key={field.key}
                                    className={`p-3.5 rounded-lg border transition-colors ${
                                      isFocused || isEditingThisField
                                        ? 'border-blue-600 bg-blue-50/30'
                                        : 'border-slate-200 bg-slate-50/60'
                                    }`}
                                  >
                                    <div className="flex items-center justify-between gap-2 mb-1.5">
                                      <label
                                        htmlFor={`mailings-input-${field.key}`}
                                        className="text-xs font-semibold text-slate-900 flex flex-wrap items-center gap-1.5 min-w-0"
                                      >
                                        <span>{field.label}</span>
                                        <span className="font-mono-tabular text-[11px] font-medium text-blue-700">
                                          {formattedToken}
                                        </span>
                                      </label>

                                      <div className="flex items-center gap-2 shrink-0">
                                        <span className="text-[11px] font-mono-tabular text-slate-500 whitespace-nowrap">
                                          Dòng {field.lineIndices.map((i) => String(i).padStart(2, '0')).join(', ')}
                                        </span>
                                        <button
                                          type="button"
                                          onClick={() => {
                                            if (isEditingThisField) {
                                              setEditingTab1FieldKey(null);
                                            } else {
                                              setEditingTab1FieldKey(field.key);
                                              setTab1RenameLabel(field.label);
                                              setTab1RenameKey(field.key);
                                            }
                                          }}
                                          title="Đổi tên hiển thị, đổi mã ký hiệu trường hoặc thay đổi sang trường khác"
                                          className={`inline-flex items-center gap-1 px-2 py-0.5 text-[11px] font-medium rounded border transition-colors cursor-pointer ${
                                            isEditingThisField
                                              ? 'bg-blue-700 text-white border-blue-700'
                                              : 'bg-white text-slate-600 border-slate-200 hover:border-blue-400 hover:text-blue-700'
                                          }`}
                                        >
                                          <Pencil className="w-2.5 h-2.5" />
                                          <span>Đổi tên / Trường</span>
                                        </button>
                                        <button
                                          type="button"
                                          onClick={() => handleRemoveFieldInTab1(field.key)}
                                          title="Xóa trường này khỏi biểu mẫu"
                                          className="inline-flex items-center gap-0.5 px-1.5 py-0.5 text-[11px] font-medium text-red-600 bg-white hover:bg-red-50 border border-slate-200 hover:border-red-300 rounded transition-colors cursor-pointer"
                                        >
                                          <X className="w-3 h-3" />
                                          <span>Xóa</span>
                                        </button>
                                      </div>
                                    </div>

                                    {/* Inline Rename / Swap Field Box inside Tab 1 Card */}
                                    {isEditingThisField && (
                                      <div className="mb-2.5 p-2.5 bg-white border border-blue-300 rounded-md space-y-2">
                                        <div className="grid grid-cols-1 sm:grid-cols-2 gap-2">
                                          <div>
                                            <label className="block text-[10px] font-semibold text-slate-600 mb-0.5">
                                              Tên hiển thị của trường:
                                            </label>
                                            <input
                                              type="text"
                                              value={tab1RenameLabel}
                                              onChange={(e) => setTab1RenameLabel(e.target.value)}
                                              className="w-full px-2 py-1 text-xs bg-slate-50 border border-slate-300 rounded focus:bg-white focus:outline-none focus:border-blue-600"
                                            />
                                          </div>
                                          <div>
                                            <label className="block text-[10px] font-semibold text-slate-600 mb-0.5">
                                              Mã ký hiệu ({currentTemplateSymbol.openSymbol}Key{currentTemplateSymbol.closeSymbol}):
                                            </label>
                                            <input
                                              type="text"
                                              value={tab1RenameKey}
                                              onChange={(e) => setTab1RenameKey(e.target.value)}
                                              className="w-full px-2 py-1 text-xs font-mono-tabular bg-slate-50 border border-slate-300 rounded focus:bg-white focus:outline-none focus:border-blue-600"
                                            />
                                          </div>
                                        </div>

                                        <div className="flex flex-wrap items-center justify-between gap-2 pt-1">
                                          <div className="flex items-center gap-1 text-[11px] text-slate-600">
                                            <ArrowRightLeft className="w-3 h-3 text-blue-700" />
                                            <span>Chọn nhanh trường có sẵn:</span>
                                            <select
                                              value={tab1RenameKey}
                                              onChange={(e) => {
                                                const chosen = customMailingsFields.find(
                                                  (f) => f.key === e.target.value
                                                );
                                                if (chosen) {
                                                  setTab1RenameKey(chosen.key);
                                                  setTab1RenameLabel(chosen.label);
                                                }
                                              }}
                                              className="px-1.5 py-0.5 text-[11px] bg-slate-50 border border-slate-300 rounded text-slate-800"
                                            >
                                              {customMailingsFields.map((opt) => (
                                                <option key={opt.key} value={opt.key}>
                                                  {opt.label} ({formatFieldToken(opt.key, currentTemplateSymbol)})
                                                </option>
                                              ))}
                                            </select>
                                          </div>

                                          <div className="flex items-center gap-1.5">
                                            <button
                                              type="button"
                                              onClick={() => setEditingTab1FieldKey(null)}
                                              className="px-2 py-1 text-[11px] font-medium text-slate-600 bg-slate-100 hover:bg-slate-200 rounded cursor-pointer"
                                            >
                                              Hủy
                                            </button>
                                            <button
                                              type="button"
                                              onClick={() =>
                                                handleRenameOrSwapFieldInTab1(
                                                  field.key,
                                                  tab1RenameLabel,
                                                  tab1RenameKey
                                                )
                                              }
                                              className="px-2.5 py-1 text-[11px] font-semibold text-white bg-blue-700 hover:bg-blue-800 rounded cursor-pointer"
                                            >
                                              Lưu đổi trường
                                            </button>
                                          </div>
                                        </div>
                                      </div>
                                    )}

                                    <input
                                      id={`mailings-input-${field.key}`}
                                      ref={(el) => {
                                        mailingsInputRefs.current[field.key] = el;
                                      }}
                                      type="text"
                                      placeholder={field.placeholder}
                                      value={val}
                                      onFocus={() => {
                                        setActiveFieldKey(field.key);
                                        if (field.lineIndices.length > 0) {
                                          setActiveFieldIndex(field.lineIndices[0]);
                                        }
                                      }}
                                      onBlur={() => {
                                        setActiveFieldKey(null);
                                        setActiveFieldIndex(null);
                                      }}
                                      onChange={(e) =>
                                        handleUpdateMailingsFieldValue(field.key, e.target.value)
                                      }
                                      className="w-full px-3 py-2 text-sm bg-white border border-slate-300 rounded-md text-slate-900 focus:outline-none focus:border-blue-600 transition-colors"
                                    />
                                  </div>
                                );
                              })}
                            </div>
                          </div>
                        )
                      ) : (
                        <div className="space-y-4 max-h-[520px] overflow-y-auto pr-1">
                          {(fillMode === 'full_lines'
                            ? currentTemplateInfo.selected_indices.map((s) => Number(s))
                            : parsedTemplateParagraphs
                                .filter((p) => {
                                  if (p.isEmpty && userReplacements[p.index] === undefined) return false;
                                  if (useSectionFilter === 'all') return true;
                                  if (useSectionFilter === 'footnote') {
                                    return p.section === 'footnote' || p.section === 'endnote';
                                  }
                                  return p.section === useSectionFilter;
                                })
                                .map((p) => p.index)
                          ).map((idx) => {
                            const idxStr = String(idx);
                            const pItem = parsedTemplateParagraphs[idx];
                            const originalText = pItem ? pItem.text : 'Dòng không tồn tại';

                            const currentVal =
                              userReplacements[idx] !== undefined
                                ? userReplacements[idx]
                                : originalText;

                            const isChanged = currentVal !== originalText;
                            const customLabel = currentTemplateInfo.custom_labels?.[idxStr];
                            const pTpl = currentTemplateInfo.paragraph_templates?.[idxStr];

                            return (
                              <div
                                key={idxStr}
                                className={`p-3.5 rounded-lg border transition-colors ${
                                  activeFieldIndex === idx
                                    ? 'border-blue-600 bg-blue-50/25'
                                    : isChanged
                                    ? 'border-emerald-500/60 bg-emerald-50/20'
                                    : 'border-slate-200 bg-slate-50/60'
                                }`}
                              >
                                <div className="flex items-start justify-between gap-2 mb-1.5">
                                  <label
                                    htmlFor={`use-txt-${idx}`}
                                    className="block text-xs font-medium text-slate-800 leading-snug"
                                  >
                                    <span className="font-mono-tabular font-semibold text-blue-800">
                                      [Dòng {String(idx).padStart(2, '0')}]
                                    </span>{' '}
                                    {pItem && pItem.section !== 'body' && (
                                      <span className="font-semibold text-blue-700 mr-1">
                                        [{pItem.sectionLabel}]
                                      </span>
                                    )}
                                    {pItem && pItem.tablePosition && (
                                      <span className="font-mono-tabular text-slate-600 mr-1">
                                        [{pItem.tablePosition}]
                                      </span>
                                    )}
                                    {customLabel ? (
                                      <span className="font-semibold text-slate-900">
                                        {customLabel}
                                      </span>
                                    ) : (
                                      <span>
                                        {originalText.length > 55
                                          ? `${originalText.slice(0, 55)}...`
                                          : originalText}
                                      </span>
                                    )}
                                  </label>

                                  {isChanged && (
                                    <button
                                      type="button"
                                      onClick={() =>
                                        setUserReplacements((prev) => ({
                                          ...prev,
                                          [idx]: originalText,
                                        }))
                                      }
                                      className="text-[11px] font-medium text-slate-500 hover:text-slate-900 whitespace-nowrap shrink-0 cursor-pointer"
                                    >
                                      Hoàn tác
                                    </button>
                                  )}
                                </div>

                                {pTpl && (
                                  <p className="text-[11px] font-mono-tabular text-blue-700 mb-1.5 truncate">
                                    Mẫu câu: {pTpl}
                                  </p>
                                )}

                                <textarea
                                  id={`use-txt-${idx}`}
                                  ref={(el) => {
                                    fieldInputRefs.current[idx] = el;
                                  }}
                                  rows={2}
                                  value={currentVal}
                                  onFocus={() => setActiveFieldIndex(idx)}
                                  onBlur={() => setActiveFieldIndex(null)}
                                  onChange={(e) => {
                                    const val = e.target.value;
                                    setUserReplacements((prev) => ({
                                      ...prev,
                                      [idx]: val,
                                    }));
                                  }}
                                  className="w-full px-3 py-2 text-sm bg-white border border-slate-300 rounded-md text-slate-900 focus:outline-none focus:border-blue-600 transition-colors"
                                />
                              </div>
                            );
                          })}
                        </div>
                      )}

                      {/* Export & Google Drive Save Action Area */}
                      <div className="pt-4 border-t border-slate-200 space-y-2.5">
                        <button
                          type="button"
                          disabled={isExporting}
                          onClick={handleExportFilledWord}
                          className="w-full flex items-center justify-center gap-2 px-4 py-3 text-sm font-semibold text-white bg-blue-700 hover:bg-blue-800 disabled:opacity-60 rounded-lg transition-colors whitespace-nowrap shrink-0 cursor-pointer"
                        >
                          <Download className="w-4 h-4" />
                          <span>
                            {isExporting
                              ? 'Đang trộn thư & tạo file Word...'
                              : 'Xuất File Word Hoàn Chỉnh (.docx)'}
                          </span>
                        </button>

                        <button
                          type="button"
                          disabled={isSavingToDrive}
                          onClick={handleSaveFilledWordToDrive}
                          className="w-full flex items-center justify-center gap-2 px-4 py-2.5 text-xs font-semibold text-slate-800 bg-slate-100 border border-slate-300 hover:bg-slate-200 disabled:opacity-60 rounded-lg transition-colors whitespace-nowrap shrink-0 cursor-pointer"
                        >
                          <CloudUpload className="w-4 h-4 text-blue-700" />
                          <span>
                            {isSavingToDrive
                              ? 'Đang tải file lên Google Drive...'
                              : 'Lưu File Word Đã Điền Lên Google Drive'}
                          </span>
                        </button>

                        {lastExportedBlob && lastExportedBlob.templateName === chosenTemplate && (
                          <button
                            type="button"
                            onClick={() =>
                              triggerBlobDownload(lastExportedBlob.blob, lastExportedBlob.fileName)
                            }
                            className="w-full flex items-center justify-center gap-2 px-4 py-2 text-xs font-semibold text-emerald-800 bg-emerald-50 border border-emerald-300 hover:bg-emerald-100 rounded-lg transition-colors whitespace-nowrap shrink-0 cursor-pointer"
                          >
                            <Check className="w-4 h-4 text-emerald-700" />
                            <span>Tải xuống lại: {lastExportedBlob.fileName}</span>
                          </button>
                        )}

                        {lastDriveSavedFile && (
                          <div className="p-3 bg-blue-50 border border-blue-200 rounded-lg flex items-center justify-between gap-2 text-xs text-blue-900">
                            <span className="font-medium truncate">
                              Đã lưu trên Drive: {lastDriveSavedFile.name}
                            </span>
                            {lastDriveSavedFile.webViewLink && (
                              <a
                                href={lastDriveSavedFile.webViewLink}
                                target="_blank"
                                rel="noopener noreferrer"
                                className="inline-flex items-center gap-1 font-semibold text-blue-700 hover:underline whitespace-nowrap shrink-0"
                              >
                                <span>Mở trên Drive</span>
                                <ExternalLink className="w-3.5 h-3.5" />
                              </a>
                            )}
                          </div>
                        )}

                        {useStatusMessage && (
                          <div
                            className={`p-3 rounded-lg text-xs font-medium border ${
                              useStatusMessage.type === 'success'
                                ? 'bg-emerald-50 text-emerald-900 border-emerald-200'
                                : useStatusMessage.type === 'warning'
                                ? 'bg-amber-50 text-amber-900 border-amber-200'
                                : 'bg-red-50 text-red-900 border-red-200'
                            }`}
                          >
                            {useStatusMessage.text}
                          </div>
                        )}
                      </div>
                    </section>
                  )}
                </div>

                {/* Right Column (7 cols on lg): Live Interactive A4 Document Preview */}
                <div className="lg:col-span-7 space-y-3">
                  <div className="flex items-center justify-between px-1">
                    <span className="text-xs font-medium text-slate-600">
                      Nhấp vào bất kỳ «Trường» hoặc dòng nào trên trang A4 để nhảy tới ô nhập liệu tương ứng
                    </span>
                    <label className="inline-flex items-center gap-1.5 text-xs text-slate-600 cursor-pointer select-none">
                      <input
                        type="checkbox"
                        checked={showLineNumbersPreview}
                        onChange={(e) => setShowLineNumbersPreview(e.target.checked)}
                        className="rounded border-slate-300 text-blue-700 focus:ring-blue-600"
                      />
                      <span>Hiện chỉ số [Dòng]</span>
                    </label>
                  </div>

                  <WordPreviewSheet
                    templateName={chosenTemplate}
                    base64Docx={currentTemplateInfo?.base64_docx}
                    paragraphs={parsedTemplateParagraphs}
                    selectedIndices={currentTemplateInfo?.selected_indices || []}
                    paragraphTemplates={currentTemplateInfo?.paragraph_templates || {}}
                    fieldValues={mailingsFieldValues}
                    replacements={userReplacements}
                    customLabels={currentTemplateInfo?.custom_labels || {}}
                    customMailingsFields={customMailingsFields}
                    fieldSymbol={currentTemplateSymbol}
                    activeIndex={activeFieldIndex}
                    activeFieldKey={activeFieldKey}
                    mode="fill"
                    showLineNumbers={showLineNumbersPreview}
                    onSelectFieldKey={(fieldKey, pIndex) => {
                      setFillMode('mailings_fields');
                      setActiveFieldKey(fieldKey);
                      setActiveFieldIndex(pIndex);
                      setTimeout(() => {
                        const el = mailingsInputRefs.current[fieldKey];
                        if (el) {
                          el.focus();
                          el.scrollIntoView({ behavior: 'smooth', block: 'center' });
                        }
                      }, 30);
                    }}
                    onSelectParagraph={(idx) => {
                      setActiveFieldIndex(idx);
                      const idxStr = String(idx);
                      const isSelectedIdx = currentTemplateInfo?.selected_indices.includes(idxStr);
                      const pTpl = currentTemplateInfo?.paragraph_templates?.[idxStr];
                      const keys = pTpl ? extractMergeFieldKeysFromText(pTpl) : [];
                      if (fillMode === 'mailings_fields' && keys.length > 0) {
                        const firstKey = keys[0];
                        setActiveFieldKey(firstKey);
                        const el = mailingsInputRefs.current[firstKey];
                        if (el) {
                          el.focus();
                          el.scrollIntoView({ behavior: 'smooth', block: 'center' });
                        }
                      } else {
                        if (!isSelectedIdx || fillMode === 'mailings_fields') {
                          setUseSectionFilter('all');
                          setFillMode(isSelectedIdx ? 'full_lines' : 'all_document_lines');
                        }
                        setTimeout(() => {
                          const inputEl = fieldInputRefs.current[idx];
                          if (inputEl) {
                            inputEl.focus();
                            inputEl.scrollIntoView({ behavior: 'smooth', block: 'center' });
                          }
                        }, 40);
                      }
                    }}
                  />
                </div>
              </div>
            )}
          </div>
        )}

        {/* ================= TAB 2: THÊM / CẤU HÌNH MẪU MỚI ================= */}
        {activeSection === 'configure' && (
          <div className="grid grid-cols-1 lg:grid-cols-12 gap-8 items-start">
            <div
              ref={configureLeftPanelRef}
              className="lg:col-span-6 lg:sticky lg:top-20 lg:max-h-[calc(100vh-5.5rem)] lg:overflow-y-auto lg:pr-1.5 space-y-6"
            >
              <section className="p-5 bg-white border border-slate-200 rounded-lg space-y-4">
                <div className="flex items-center justify-between gap-2">
                  <div>
                    <h2 className="text-base font-semibold text-slate-900">
                      {editingExistingTemplateName
                        ? `Cấu hình vị trí trường Mailings: ${editingExistingTemplateName}`
                        : 'Tải lên và chọn vị trí chèn trường dữ liệu (Mailings Word)'}
                    </h2>
                    <p className="text-xs text-slate-500 mt-0.5">
                      Chọn dòng trong Word, đặt con trỏ hoặc bôi đen vị trí cần thay thế rồi bấm vào các trường mặc định «Ho_Va_Ten», «Ngay», «So_Tien_So»...
                    </p>
                  </div>
                  {editingExistingTemplateName && (
                    <div className="flex items-center gap-3 shrink-0">
                      <button
                        type="button"
                        onClick={() => handleOpenDuplicateModal(editingExistingTemplateName)}
                        className="inline-flex items-center gap-1 text-xs font-semibold text-slate-700 hover:text-blue-700 whitespace-nowrap cursor-pointer"
                      >
                        <CopyPlus className="w-3.5 h-3.5" />
                        <span>Nhân bản mẫu này</span>
                      </button>
                      <button
                        type="button"
                        onClick={() => {
                          setEditingExistingTemplateName(null);
                          setTemplateNameInput('');
                          setUploadedFileName('');
                          setUploadedBase64('');
                          setUploadedAllParagraphs([]);
                          setUploadedNonEmptyParagraphs([]);
                          setChosenIndices([]);
                          setParagraphTemplatesInput({});
                          setDetectedDefaultValues({});
                          setCustomLabelsInput({});
                          setActiveConfigParagraphIndex(null);
                          setConfigStatusMessage(null);
                        }}
                        className="text-xs font-semibold text-blue-700 hover:text-blue-800 whitespace-nowrap cursor-pointer"
                      >
                        + Tạo mẫu mới
                      </button>
                    </div>
                  )}
                </div>

                <div className="grid grid-cols-1 sm:grid-cols-3 gap-4">
                  <div className="sm:col-span-2">
                    <label
                      htmlFor="template-name-input"
                      className="block text-xs font-medium text-slate-700 mb-1.5"
                    >
                      Đặt tên gợi nhớ cho mẫu (Ví dụ: Giấy đề nghị A):
                    </label>
                    <input
                      id="template-name-input"
                      type="text"
                      placeholder="Nhập tên biểu mẫu..."
                      value={templateNameInput}
                      onChange={(e) => setTemplateNameInput(e.target.value)}
                      className="w-full px-3.5 py-2 text-sm bg-white border border-slate-300 rounded-lg text-slate-900 focus:outline-none focus:border-blue-600"
                    />
                  </div>
                  <div>
                    <label
                      htmlFor="template-category-input"
                      className="block text-xs font-medium text-slate-700 mb-1.5"
                    >
                      Phân nhóm:
                    </label>
                    <input
                      id="template-category-input"
                      type="text"
                      placeholder="Hành chính, Kế toán..."
                      value={templateCategoryInput}
                      onChange={(e) => setTemplateCategoryInput(e.target.value)}
                      className="w-full px-3.5 py-2 text-sm bg-white border border-slate-300 rounded-lg text-slate-900 focus:outline-none focus:border-blue-600"
                    />
                  </div>
                </div>

                {/* File Upload Dropzone + Google Drive Picker Button */}
                <div>
                  <div className="flex flex-wrap items-center justify-between gap-2 mb-1.5">
                    <span className="text-xs font-medium text-slate-700">
                      Tải file Word mẫu (.docx):
                    </span>
                    <div className="flex items-center gap-3">
                      <button
                        type="button"
                        onClick={() => setIsDriveModalOpen(true)}
                        className="inline-flex items-center gap-1 text-xs font-semibold text-blue-700 hover:text-blue-800 cursor-pointer"
                      >
                        <Cloud className="w-3.5 h-3.5" />
                        <span>Chọn từ Google Drive</span>
                      </button>
                      <span aria-hidden="true" className="text-slate-300">·</span>
                      <button
                        type="button"
                        onClick={handleLoadTestUploadDocx}
                        className="inline-flex items-center gap-1 text-xs font-semibold text-slate-600 hover:text-slate-900 cursor-pointer"
                      >
                        <FileSpreadsheet className="w-3.5 h-3.5" />
                        <span>Nạp file .docx thử nghiệm</span>
                      </button>
                    </div>
                  </div>

                  <div
                    onDragOver={(e) => e.preventDefault()}
                    onDrop={(e) => {
                      e.preventDefault();
                      const file = e.dataTransfer.files?.[0];
                      if (file) handleFileUpload(file);
                    }}
                    onClick={() => fileInputRef.current?.click()}
                    className={`border border-dashed border-slate-300 hover:border-blue-600 bg-slate-50/70 hover:bg-blue-50/20 rounded-lg cursor-pointer transition-colors ${
                      uploadedFileName
                        ? 'px-3.5 py-2.5 flex items-center justify-between gap-3'
                        : 'p-4 text-center'
                    }`}
                  >
                    <input
                      ref={fileInputRef}
                      type="file"
                      accept=".docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                      className="hidden"
                      onChange={(e) => {
                        const file = e.target.files?.[0];
                        if (file) handleFileUpload(file);
                      }}
                    />
                    {uploadedFileName ? (
                      <>
                        <div className="flex items-center gap-2.5 min-w-0">
                          <Upload className="w-4 h-4 text-blue-700 shrink-0" />
                          <div className="min-w-0">
                            <p className="text-xs font-semibold text-slate-900 truncate">
                              {uploadedFileName}
                            </p>
                            <p className="text-[11px] text-slate-500 font-mono-tabular">
                              {uploadedNonEmptyParagraphs.length} đoạn văn bản
                            </p>
                          </div>
                        </div>
                        <span className="text-xs font-semibold text-blue-700 whitespace-nowrap shrink-0">
                          Đổi file .docx khác
                        </span>
                      </>
                    ) : (
                      <>
                        <Upload className="w-5 h-5 text-slate-500 mx-auto mb-1.5" />
                        <div className="space-y-0.5">
                          <p className="text-sm font-medium text-slate-800">
                            Nhấp để chọn từ máy tính hoặc kéo thả file Word (.docx) vào đây
                          </p>
                          <p className="text-xs text-slate-500">
                            Hoặc bấm &ldquo;Chọn từ Google Drive&rdquo; để lấy file .docx / Google Docs trên đám mây
                          </p>
                        </div>
                      </>
                    )}
                  </div>
                </div>

                {configStatusMessage && (
                  <div
                    className={`p-3.5 rounded-lg text-xs font-medium border flex items-center justify-between gap-3 ${
                      configStatusMessage.type === 'success'
                        ? 'bg-emerald-50 text-emerald-900 border-emerald-200'
                        : configStatusMessage.type === 'warning'
                        ? 'bg-amber-50 text-amber-900 border-amber-200'
                        : 'bg-red-50 text-red-900 border-red-200'
                    }`}
                  >
                    <span>{configStatusMessage.text}</span>
                    {configStatusMessage.type === 'success' &&
                      chosenTemplate === templateNameInput.trim() && (
                        <button
                          type="button"
                          onClick={() => setActiveSection('use')}
                          className="px-3 py-1.5 text-xs font-semibold text-white bg-emerald-700 hover:bg-emerald-800 rounded-md whitespace-nowrap shrink-0 cursor-pointer"
                        >
                          Mở điền mẫu ngay
                        </button>
                      )}
                  </div>
                )}

                {uploadedNonEmptyParagraphs.length > 0 && (
                  <div ref={mailingsEditorAnchorRef} className="pt-1 space-y-5">
                    <MailingsPositionEditor
                      activeParagraph={activeConfigParagraph}
                      currentTemplateString={activeConfigTemplateString}
                      customMailingsFields={customMailingsFields}
                      fieldSymbol={fieldSymbol}
                      onUpdateFieldSymbol={handleUpdateFieldSymbolInConfig}
                      onUpdateParagraphTemplate={handleUpdateParagraphTemplateInConfig}
                      onResetParagraphTemplate={handleResetParagraphTemplateInConfig}
                      onAddCustomField={handleAddCustomMailingsField}
                      onRenameOrUpdateField={handleRenameOrUpdateMailingsField}
                      onDeleteField={handleDeleteMailingsField}
                      onResetDefaultPalette={handleResetDefaultMailingsPalette}
                    />

                    <div className="space-y-3 pt-2 border-t border-slate-200">
                      <div className="flex flex-wrap items-center justify-between gap-2">
                        <div>
                          <h3 className="text-sm font-semibold text-slate-900">
                            Danh sách dòng văn bản trong Word (Nhấp vào dòng để đặt vị trí chèn trường)
                          </h3>
                          <p className="text-xs text-slate-500 font-mono-tabular mt-0.5">
                            Đã đánh dấu {chosenIndices.length} / {uploadedNonEmptyParagraphs.length} dòng
                          </p>
                        </div>

                        <div className="flex items-center gap-2">
                          <button
                            type="button"
                            onClick={() =>
                              setChosenIndices(
                                uploadedNonEmptyParagraphs.map((item) => String(item.index))
                              )
                            }
                            className="px-2.5 py-1 text-xs font-medium text-slate-700 bg-slate-100 hover:bg-slate-200 rounded-md whitespace-nowrap shrink-0 cursor-pointer"
                          >
                            Chọn tất cả
                          </button>
                          <button
                            type="button"
                            onClick={() => {
                              setChosenIndices([]);
                              setParagraphTemplatesInput({});
                            }}
                            className="px-2.5 py-1 text-xs font-medium text-slate-700 bg-slate-100 hover:bg-slate-200 rounded-md whitespace-nowrap shrink-0 cursor-pointer"
                          >
                            Xóa chọn hết
                          </button>
                        </div>
                      </div>

                      {/* Section Filter Bar for Headnote, Body/Table, Footnote, Footer */}
                      <div className="flex flex-wrap items-center gap-1">
                        {[
                          { id: 'all', label: `Tất cả (${uploadedNonEmptyParagraphs.length})` },
                          {
                            id: 'header',
                            label: `Đầu trang / Headnote (${
                              uploadedNonEmptyParagraphs.filter((p) => p.section === 'header').length
                            })`,
                          },
                          {
                            id: 'body',
                            label: `Nội dung & Bảng (${
                              uploadedNonEmptyParagraphs.filter((p) => p.section === 'body').length
                            })`,
                          },
                          {
                            id: 'footnote',
                            label: `Chú thích / Footnote (${
                              uploadedNonEmptyParagraphs.filter(
                                (p) => p.section === 'footnote' || p.section === 'endnote'
                              ).length
                            })`,
                          },
                          {
                            id: 'footer',
                            label: `Chân trang / Footer (${
                              uploadedNonEmptyParagraphs.filter((p) => p.section === 'footer').length
                            })`,
                          },
                        ].map((sec) => (
                          <button
                            key={sec.id}
                            type="button"
                            onClick={() => setConfigureSectionFilter(sec.id as any)}
                            className={`px-2.5 py-1 text-[11px] font-medium rounded-md transition-colors cursor-pointer ${
                              configureSectionFilter === sec.id
                                ? 'bg-slate-900 text-white'
                                : 'bg-slate-100 text-slate-600 hover:text-slate-900'
                            }`}
                          >
                            {sec.label}
                          </button>
                        ))}
                      </div>

                      <div className="relative">
                        <Search className="w-4 h-4 text-slate-400 absolute left-3 top-1/2 -translate-y-1/2" />
                        <input
                          type="text"
                          placeholder="Tìm kiếm dòng theo nội dung, vùng (Headnote, Footnote, Bảng) hoặc mã «Trường»..."
                          value={paragraphSearchQuery}
                          onChange={(e) => setParagraphSearchQuery(e.target.value)}
                          className="w-full pl-9 pr-3 py-2 text-xs bg-slate-50 border border-slate-200 rounded-lg text-slate-900 focus:outline-none focus:bg-white focus:border-blue-600"
                        />
                      </div>

                      <div className="space-y-2 max-h-[360px] overflow-y-auto pr-1">
                        {filteredConfigureParagraphs.map((item) => {
                          const idxStr = String(item.index);
                          const isChecked = chosenIndices.includes(idxStr);
                          const isCurrentTarget = activeConfigParagraphIndex === item.index;
                          const tplText = paragraphTemplatesInput[idxStr] ?? item.text;
                          const embeddedKeys = extractMergeFieldKeysFromText(tplText, fieldSymbol);

                          return (
                            <div
                              key={item.index}
                              onClick={() => setActiveConfigParagraphIndex(item.index)}
                              className={`p-3 rounded-lg border transition-colors cursor-pointer ${
                                isCurrentTarget
                                  ? 'bg-blue-50/70 border-blue-600'
                                  : isChecked
                                  ? 'bg-amber-50/40 border-amber-400/80 hover:border-amber-500'
                                  : 'bg-white border-slate-200 hover:border-slate-300'
                              }`}
                            >
                              <div className="flex items-start justify-between gap-3">
                                <label
                                  onClick={(e) => e.stopPropagation()}
                                  className="flex items-start gap-2.5 cursor-pointer select-none flex-1 min-w-0"
                                >
                                  <input
                                    type="checkbox"
                                    checked={isChecked}
                                    onChange={() => handleToggleConfigureIndex(item.index)}
                                    className="mt-0.5 rounded border-slate-300 text-blue-700 focus:ring-blue-600"
                                  />
                                  <div className="text-xs leading-relaxed text-slate-800 min-w-0 flex-1">
                                    <span className="font-mono-tabular font-semibold text-blue-800">
                                      [Dòng {String(item.index).padStart(2, '0')}]
                                    </span>{' '}
                                    {item.section !== 'body' && (
                                      <span className="font-semibold text-blue-700 mr-1">
                                        [{item.sectionLabel}]
                                      </span>
                                    )}
                                    {item.tablePosition && (
                                      <span className="font-mono-tabular text-slate-600 mr-1">
                                        [{item.tablePosition}]
                                      </span>
                                    )}
                                    <span className="break-words font-medium">{tplText}</span>
                                  </div>
                                </label>

                                <div className="flex items-center gap-1.5 shrink-0">
                                  {(isChecked || paragraphTemplatesInput[idxStr] !== undefined) && (
                                    <button
                                      type="button"
                                      onClick={(e) => {
                                        e.stopPropagation();
                                        handleResetParagraphTemplateInConfig(item.index);
                                      }}
                                      title="Xóa chọn và khôi phục dòng này"
                                      className="inline-flex items-center gap-1 px-2 py-1 text-[11px] font-medium text-red-600 bg-red-50 hover:bg-red-100 rounded border border-red-200 whitespace-nowrap cursor-pointer"
                                    >
                                      <X className="w-3 h-3" />
                                      <span>Xóa chọn</span>
                                    </button>
                                  )}
                                  <button
                                    type="button"
                                    onClick={(e) => {
                                      e.stopPropagation();
                                      setActiveConfigParagraphIndex(item.index);
                                    }}
                                    className={`inline-flex items-center gap-1 px-2 py-1 text-[11px] font-medium rounded whitespace-nowrap shrink-0 cursor-pointer ${
                                      isCurrentTarget
                                        ? 'bg-blue-700 text-white'
                                        : 'bg-slate-100 text-slate-700 hover:bg-slate-200'
                                    }`}
                                  >
                                    <MousePointerClick className="w-3 h-3" />
                                    <span>{isCurrentTarget ? 'Đang chọn vị trí' : 'Đặt trường'}</span>
                                  </button>
                                </div>
                              </div>

                              {embeddedKeys.length > 0 && (
                                <div className="mt-1.5 pl-6 flex flex-wrap items-center gap-1.5 text-[11px] text-blue-800">
                                  <Tag className="w-3 h-3 text-blue-600 shrink-0" />
                                  <span>Trường đã gắn tại dòng này:</span>
                                  {embeddedKeys.map((k) => (
                                    <span
                                      key={k}
                                      className="inline-flex items-center gap-1 px-1.5 py-0.5 bg-white border border-blue-200 rounded font-mono-tabular font-semibold"
                                    >
                                      <span>{formatFieldToken(k, fieldSymbol)}</span>
                                      <button
                                        type="button"
                                        onClick={(e) => {
                                          e.stopPropagation();
                                          const cleaned = removeFieldKeyFromTemplateString(
                                            tplText,
                                            k,
                                            fieldSymbol
                                          );
                                          if (
                                            cleaned === item.text &&
                                            extractMergeFieldKeysFromText(cleaned, fieldSymbol)
                                              .length === 0
                                          ) {
                                            handleResetParagraphTemplateInConfig(item.index);
                                          } else {
                                            handleUpdateParagraphTemplateInConfig(
                                              item.index,
                                              cleaned
                                            );
                                          }
                                        }}
                                        title={`Xóa trường ${formatFieldToken(k, fieldSymbol)} khỏi dòng này`}
                                        className="text-slate-400 hover:text-red-600 cursor-pointer"
                                      >
                                        <X className="w-3 h-3" />
                                      </button>
                                    </span>
                                  ))}
                                </div>
                              )}
                            </div>
                          );
                        })}
                      </div>

                      <div className="sticky bottom-0 bg-white pt-3 pb-1 border-t border-slate-200 z-10">
                        <button
                          type="button"
                          onClick={handleSaveTemplateToLibrary}
                          className="w-full flex items-center justify-center gap-2 px-4 py-3 text-sm font-semibold text-white bg-blue-700 hover:bg-blue-800 rounded-lg transition-colors whitespace-nowrap shrink-0 cursor-pointer"
                        >
                          <Check className="w-4 h-4" />
                          <span>Lưu Mẫu & Cấu Hình Trường Mailings Vào Thư Viện</span>
                        </button>
                      </div>
                    </div>
                  </div>
                )}
              </section>
            </div>

            <div className="lg:col-span-6 space-y-2">
              <div className="px-1 text-xs font-medium text-slate-600">
                Nhấp vào bất kỳ đoạn văn nào trên trang A4 bên dưới để chọn dòng đó và đặt vị trí chèn «Trường Mailings»
              </div>
              <WordPreviewSheet
                templateName={templateNameInput || uploadedFileName || 'Xem trước vị trí trường Mailings'}
                base64Docx={uploadedBase64}
                paragraphs={uploadedAllParagraphs}
                selectedIndices={chosenIndices}
                paragraphTemplates={paragraphTemplatesInput}
                fieldValues={detectedDefaultValues}
                customLabels={customLabelsInput}
                customMailingsFields={customMailingsFields}
                fieldSymbol={fieldSymbol}
                activeIndex={activeConfigParagraphIndex}
                mode="configure"
                onSelectParagraph={(idx) => {
                  setActiveConfigParagraphIndex(idx);
                  if (configureLeftPanelRef.current && mailingsEditorAnchorRef.current) {
                    const panelEl = configureLeftPanelRef.current;
                    const anchorEl = mailingsEditorAnchorRef.current;
                    const targetTop = Math.max(0, anchorEl.offsetTop - 16);
                    if (panelEl.scrollTop > targetTop + 80) {
                      panelEl.scrollTo({ top: targetTop, behavior: 'smooth' });
                    }
                  }
                }}
              />
            </div>
          </div>
        )}

        {/* ================= TAB 3: KHO BIỂU MẪU & LỊCH SỬ XUẤT FILE (LANDING PAGE) ================= */}
        {activeSection === 'library' && (
          <div className="space-y-8">
            {libraryBannerMessage && (
              <div className="p-3.5 bg-emerald-50 border border-emerald-200 rounded-lg flex items-center justify-between gap-3 text-xs font-medium text-emerald-900">
                <span>{libraryBannerMessage}</span>
                <button
                  type="button"
                  onClick={() => setLibraryBannerMessage(null)}
                  className="text-emerald-700 hover:text-emerald-950 cursor-pointer"
                >
                  <X className="w-4 h-4" />
                </button>
              </div>
            )}

            {/* Google Drive Shared Cloud Storage Banner */}
            <div className="p-5 bg-white border border-slate-200 rounded-lg flex flex-col lg:flex-row lg:items-center justify-between gap-4">
              <div className="flex items-start gap-3.5">
                <div className="w-10 h-10 rounded-lg bg-blue-50 border border-blue-200 flex items-center justify-center shrink-0">
                  <Globe className="w-5 h-5 text-blue-700" />
                </div>
                <div className="space-y-1">
                  <div className="flex flex-wrap items-center gap-2">
                    <h3 className="text-sm font-bold text-slate-900">
                      {sharedDriveInfo
                        ? 'Kho Biểu Mẫu Đang Lưu Trữ & Chia Sẻ Trực Tiếp Trên Google Drive'
                        : 'Lưu Trữ Kho Mẫu Lên Google Drive & Chia Sẻ Cho Mọi Người Truy Cập App'}
                    </h3>
                    {sharedDriveInfo && (
                      <span className="text-[11px] font-mono-tabular font-semibold text-emerald-700">
                        ● Đã bật chia sẻ công khai trên Drive
                      </span>
                    )}
                  </div>

                  {sharedDriveInfo ? (
                    <div className="text-xs text-slate-600 flex flex-wrap items-center gap-x-2 gap-y-1">
                      <span>
                        Thư mục Drive:{' '}
                        <strong className="text-slate-900">
                          {sharedDriveInfo.folderName || SHARED_DRIVE_FOLDER_NAME}
                        </strong>
                      </span>
                      {sharedDriveInfo.ownerEmail && (
                        <>
                          <span aria-hidden="true">·</span>
                          <span>
                            Tài khoản lưu trữ: <strong className="text-slate-900">{sharedDriveInfo.ownerEmail}</strong>
                          </span>
                        </>
                      )}
                      <span aria-hidden="true">·</span>
                      <span className="font-mono-tabular">
                        Cập nhật: {new Date(sharedDriveInfo.updatedAt).toLocaleString('vi-VN')}
                      </span>
                      {sharedDriveInfo.folderWebViewLink && (
                        <>
                          <span aria-hidden="true">·</span>
                          <a
                            href={sharedDriveInfo.folderWebViewLink}
                            target="_blank"
                            rel="noopener noreferrer"
                            className="inline-flex items-center gap-1 font-semibold text-blue-700 hover:underline"
                          >
                            <span>Mở thư mục trên Google Drive</span>
                            <ExternalLink className="w-3 h-3" />
                          </a>
                        </>
                      )}
                    </div>
                  ) : (
                    <p className="text-xs text-slate-600 leading-relaxed">
                      Nhấn nút <strong>&ldquo;Lưu & Chia Sẻ Kho Mẫu Lên Google Drive&rdquo;</strong> để tải toàn bộ các biểu mẫu (.docx & cấu hình trường Mailings) vào Google Drive của bạn và tự động cấp quyền đọc cho bất kỳ ai mở link ứng dụng.
                    </p>
                  )}
                </div>
              </div>

              <div className="flex flex-wrap items-center gap-2 shrink-0">
                <button
                  type="button"
                  disabled={isPublishingSharedDrive}
                  onClick={handleRequestPublishToSharedDrive}
                  className="inline-flex items-center gap-1.5 px-3.5 py-2 text-xs font-semibold text-white bg-blue-700 hover:bg-blue-800 disabled:opacity-60 rounded-lg transition-colors whitespace-nowrap shrink-0 cursor-pointer"
                >
                  <CloudUpload className="w-3.5 h-3.5" />
                  <span>
                    {isPublishingSharedDrive
                      ? 'Đang đồng bộ lên Google Drive...'
                      : sharedDriveInfo
                      ? 'Cập nhật Kho Mẫu lên Google Drive'
                      : 'Lưu & Chia Sẻ Kho Mẫu Lên Google Drive'}
                  </span>
                </button>

                <button
                  type="button"
                  onClick={handleShareLibraryLink}
                  className="inline-flex items-center gap-1.5 px-3.5 py-2 text-xs font-semibold text-emerald-800 bg-emerald-50 border border-emerald-300 hover:bg-emerald-100 rounded-lg transition-colors whitespace-nowrap shrink-0 cursor-pointer"
                >
                  <Share2 className="w-3.5 h-3.5 text-emerald-700" />
                  <span>Copy Link App (Đọc từ Drive)</span>
                </button>
              </div>
            </div>

            {/* Saved Templates High-Density Table with Search, Category Filter, Sort, and Google Drive Integration */}
            <section className="bg-white border border-slate-200 rounded-lg overflow-hidden">
              <div className="flex flex-wrap items-center justify-between gap-4 px-6 py-4 border-b border-slate-200">
                <div>
                  <h2 className="text-base font-semibold text-slate-900">
                    Danh sách Biểu Mẫu Word trong Kho Dùng Chung
                  </h2>
                  <div className="flex flex-wrap items-center gap-2 text-xs text-slate-500 mt-0.5">
                    <span className="font-mono-tabular">
                      Hiển thị {filteredLibraryTemplates.length} / {templateNames.length} biểu mẫu
                    </span>
                    <span aria-hidden="true">·</span>
                    <span>
                      {sharedDriveInfo
                        ? 'Dữ liệu được nạp trực tiếp từ Google Drive cho mọi người truy cập'
                        : 'Hỗ trợ lưu trữ trên Google Drive và chia sẻ cho mọi người dùng'}
                    </span>
                  </div>
                </div>
                <div className="flex flex-wrap items-center gap-2">
                  <button
                    type="button"
                    onClick={handleRefreshSharedLibrary}
                    title="Tải lại các mẫu mới nhất từ Google Drive dùng chung"
                    className="inline-flex items-center gap-1.5 px-3 py-2 text-xs font-medium text-slate-700 bg-slate-100 hover:bg-slate-200 rounded-lg transition-colors whitespace-nowrap shrink-0 cursor-pointer"
                  >
                    <RefreshCw className="w-3.5 h-3.5 text-slate-600" />
                    <span>Tải mới nhất từ Drive</span>
                  </button>
                  <button
                    type="button"
                    onClick={() => setIsDriveModalOpen(true)}
                    className="inline-flex items-center gap-1.5 px-3.5 py-2 text-xs font-semibold text-slate-700 bg-slate-100 hover:bg-slate-200 rounded-lg transition-colors whitespace-nowrap shrink-0 cursor-pointer"
                  >
                    <Cloud className="w-3.5 h-3.5 text-blue-700" />
                    <span>Quản lý Google Drive</span>
                  </button>
                  <button
                    type="button"
                    onClick={() => {
                      setEditingExistingTemplateName(null);
                      setActiveSection('configure');
                    }}
                    className="inline-flex items-center gap-1.5 px-3.5 py-2 text-xs font-semibold text-white bg-blue-700 hover:bg-blue-800 rounded-lg transition-colors whitespace-nowrap shrink-0 cursor-pointer"
                  >
                    <Plus className="w-3.5 h-3.5" />
                    <span>Thêm mẫu mới</span>
                  </button>
                </div>
              </div>

              {/* Filter & Search Bar for Kho Biểu Mẫu */}
              <div className="px-6 py-3.5 bg-slate-50/80 border-b border-slate-200 flex flex-col lg:flex-row lg:items-center justify-between gap-3">
                <div className="relative flex-1 max-w-md">
                  <Search className="w-4 h-4 text-slate-400 absolute left-3 top-1/2 -translate-y-1/2" />
                  <input
                    type="text"
                    value={librarySearchQuery}
                    onChange={(e) => setLibrarySearchQuery(e.target.value)}
                    placeholder="Tìm theo tên mẫu, phân nhóm, hoặc mã «Trường»..."
                    className="w-full pl-9 pr-8 py-1.5 text-xs bg-white border border-slate-300 rounded-lg text-slate-900 focus:outline-none focus:border-blue-600"
                  />
                  {librarySearchQuery && (
                    <button
                      type="button"
                      onClick={() => setLibrarySearchQuery('')}
                      className="absolute right-2.5 top-1/2 -translate-y-1/2 text-slate-400 hover:text-slate-700 cursor-pointer"
                    >
                      <X className="w-3.5 h-3.5" />
                    </button>
                  )}
                </div>

                <div className="flex flex-wrap items-center gap-2.5">
                  <div className="flex flex-wrap items-center gap-1 p-1 bg-slate-200/70 rounded-lg">
                    <button
                      type="button"
                      onClick={() => setLibraryCategoryFilter('ALL')}
                      className={`px-2.5 py-1 text-xs font-medium rounded-md transition-colors whitespace-nowrap shrink-0 cursor-pointer ${
                        libraryCategoryFilter === 'ALL'
                          ? 'bg-white text-slate-900 shadow-2xs font-semibold'
                          : 'text-slate-600 hover:text-slate-900'
                      }`}
                    >
                      Tất cả nhóm ({templateNames.length})
                    </button>
                    {availableCategories.map((cat) => (
                      <button
                        key={cat}
                        type="button"
                        onClick={() => setLibraryCategoryFilter(cat)}
                        className={`px-2.5 py-1 text-xs font-medium rounded-md transition-colors whitespace-nowrap shrink-0 cursor-pointer ${
                          libraryCategoryFilter === cat
                            ? 'bg-white text-slate-900 shadow-2xs font-semibold'
                            : 'text-slate-600 hover:text-slate-900'
                        }`}
                      >
                        {cat}
                      </button>
                    ))}
                  </div>

                  <div className="flex items-center gap-1.5">
                    <ArrowUpDown className="w-3.5 h-3.5 text-slate-500 shrink-0" />
                    <select
                      value={librarySortBy}
                      onChange={(e) => setLibrarySortBy(e.target.value as LibrarySortOption)}
                      className="px-2.5 py-1.5 text-xs bg-white border border-slate-300 rounded-lg text-slate-700 focus:outline-none focus:border-blue-600"
                    >
                      <option value="updated_desc">Sắp xếp: Mới cập nhật</option>
                      <option value="name_asc">Sắp xếp: Tên A → Z</option>
                      <option value="fields_desc">Sắp xếp: Nhiều trường nhất</option>
                      <option value="lines_desc">Sắp xếp: Nhiều dòng nhất</option>
                    </select>
                  </div>

                  {(librarySearchQuery || libraryCategoryFilter !== 'ALL') && (
                    <button
                      type="button"
                      onClick={() => {
                        setLibrarySearchQuery('');
                        setLibraryCategoryFilter('ALL');
                      }}
                      className="text-xs font-medium text-blue-700 hover:text-blue-800 whitespace-nowrap shrink-0 cursor-pointer"
                    >
                      Xóa bộ lọc
                    </button>
                  )}
                </div>
              </div>

              {templateNames.length === 0 ? (
                <div className="p-10 text-center text-sm text-slate-500">
                  Chưa có biểu mẫu nào trong thư viện.
                </div>
              ) : filteredLibraryTemplates.length === 0 ? (
                <div className="p-10 text-center space-y-2">
                  <p className="text-sm font-medium text-slate-700">
                    Không tìm thấy biểu mẫu nào khớp với bộ lọc hiện tại.
                  </p>
                  <button
                    type="button"
                    onClick={() => {
                      setLibrarySearchQuery('');
                      setLibraryCategoryFilter('ALL');
                    }}
                    className="px-3 py-1.5 text-xs font-semibold text-blue-700 bg-blue-50 hover:bg-blue-100 rounded-md cursor-pointer"
                  >
                    Đặt lại bộ lọc tìm kiếm
                  </button>
                </div>
              ) : (
                <div className="overflow-x-auto">
                  <table className="w-full text-left border-collapse">
                    <thead>
                      <tr className="border-b border-slate-200 bg-slate-50 text-xs font-semibold text-slate-600">
                        <th className="py-3 px-5">Tên biểu mẫu</th>
                        <th className="py-3 px-4">Phân nhóm</th>
                        <th className="py-3 px-4">Các trường Mailings đã gắn</th>
                        <th className="py-3 px-4 text-right">Dòng động</th>
                        <th className="py-3 px-5 text-right">Thao tác</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-200 text-sm">
                      {filteredLibraryTemplates.map(([name, item]) => {
                        const fieldKeys = item.mailings_fields
                          ? Object.keys(item.mailings_fields)
                          : [];
                        return (
                          <tr key={name} className="hover:bg-slate-50/80 transition-colors">
                            <td className="py-3 px-5">
                              <div className="font-medium text-slate-900 flex items-center gap-2">
                                <span>{name}</span>
                                {item.drive_web_view_link && (
                                  <a
                                    href={item.drive_web_view_link}
                                    target="_blank"
                                    rel="noopener noreferrer"
                                    title="Mở file Word này trên Google Drive"
                                    className="inline-flex items-center gap-0.5 text-[11px] font-medium text-blue-700 hover:underline"
                                  >
                                    <span>Drive</span>
                                    <ExternalLink className="w-3 h-3" />
                                  </a>
                                )}
                              </div>
                              <div className="text-[11px] font-mono-tabular text-slate-400">
                                {item.file_path}
                              </div>
                            </td>
                            <td className="py-3 px-4 text-xs text-slate-600">
                              {item.category || 'Biểu mẫu chung'}
                            </td>
                            <td className="py-3 px-4 text-xs font-mono-tabular text-blue-800 max-w-md truncate">
                              {fieldKeys.length > 0
                                ? `(${fieldKeys.length}) ${fieldKeys.map((k) => `«${k}»`).join(' · ')}`
                                : 'Sửa nguyên dòng'}
                            </td>
                            <td className="py-3 px-4 text-right font-mono-tabular font-medium text-slate-700">
                              {item.selected_indices.length} dòng
                            </td>
                            <td className="py-3 px-5 text-right">
                              <div className="inline-flex items-center justify-end gap-1.5">
                                <button
                                  type="button"
                                  onClick={() => {
                                    setChosenTemplate(name);
                                    setActiveSection('use');
                                  }}
                                  className="px-2.5 py-1 text-xs font-semibold text-blue-700 bg-blue-50 hover:bg-blue-100 rounded transition-colors whitespace-nowrap shrink-0 cursor-pointer"
                                >
                                  Điền mẫu
                                </button>
                                <button
                                  type="button"
                                  onClick={() => handleShareSingleTemplateLink(name)}
                                  title="Sao chép link chia sẻ trực tiếp biểu mẫu này"
                                  className="inline-flex items-center gap-1 px-2.5 py-1 text-xs font-medium text-emerald-800 bg-emerald-50 hover:bg-emerald-100 rounded transition-colors whitespace-nowrap shrink-0 cursor-pointer"
                                >
                                  <Share2 className="w-3.5 h-3.5 text-emerald-700" />
                                  <span>Share link</span>
                                </button>
                                <button
                                  type="button"
                                  onClick={() => handleOpenDuplicateModal(name)}
                                  title="Nhân bản biểu mẫu này thành một mẫu mới"
                                  className="inline-flex items-center gap-1 px-2.5 py-1 text-xs font-medium text-slate-700 bg-slate-100 hover:bg-slate-200 rounded transition-colors whitespace-nowrap shrink-0 cursor-pointer"
                                >
                                  <CopyPlus className="w-3.5 h-3.5 text-blue-700" />
                                  <span>Nhân bản</span>
                                </button>
                                <button
                                  type="button"
                                  onClick={() => handleEditExistingTemplateConfig(name)}
                                  className="px-2.5 py-1 text-xs font-medium text-slate-700 bg-slate-100 hover:bg-slate-200 rounded transition-colors whitespace-nowrap shrink-0 cursor-pointer"
                                >
                                  Sửa «Trường»
                                </button>
                                <button
                                  type="button"
                                  onClick={() =>
                                    handleQuickUploadDocxToDrive(
                                      name,
                                      {},
                                      item.file_path.split('/').pop() || `${name}_goc.docx`
                                    )
                                  }
                                  title="Tải file .docx mẫu gốc lên Google Drive"
                                  className="p-1.5 text-blue-700 hover:text-blue-900 rounded hover:bg-blue-50 cursor-pointer"
                                >
                                  <CloudUpload className="w-3.5 h-3.5" />
                                </button>
                                <button
                                  type="button"
                                  onClick={() => handleDownloadOriginalTemplate(name)}
                                  title="Tải file .docx gốc về máy"
                                  className="p-1.5 text-slate-600 hover:text-slate-900 rounded hover:bg-slate-100 cursor-pointer"
                                >
                                  <Download className="w-3.5 h-3.5" />
                                </button>
                                <button
                                  type="button"
                                  onClick={() => handleDeleteTemplate(name)}
                                  title="Xóa mẫu"
                                  className="p-1.5 text-red-600 hover:text-red-700 rounded hover:bg-red-50 cursor-pointer"
                                >
                                  <Trash2 className="w-3.5 h-3.5" />
                                </button>
                              </div>
                            </td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>
              )}
            </section>

            {/* Export History Table */}
            <section className="bg-white border border-slate-200 rounded-lg overflow-hidden">
              <div className="flex flex-wrap items-center justify-between gap-4 px-6 py-4 border-b border-slate-200">
                <div>
                  <div className="flex items-center gap-2">
                    <History className="w-4 h-4 text-slate-600" />
                    <h2 className="text-base font-semibold text-slate-900">
                      Lịch sử xuất file Word
                    </h2>
                  </div>
                  <div className="flex items-center gap-2 text-xs text-slate-500 mt-0.5">
                    <span className="font-mono-tabular">
                      Hiển thị {filteredExportHistory.length} / {exportHistory.length} lượt xuất file
                    </span>
                    <span aria-hidden="true">·</span>
                    <span>Tìm kiếm được cả nội dung đã điền bên trong tài liệu</span>
                  </div>
                </div>
                {exportHistory.length > 0 && (
                  <button
                    type="button"
                    onClick={() => {
                      clearExportHistory();
                      setExportHistory([]);
                    }}
                    className="text-xs font-medium text-slate-500 hover:text-red-600 whitespace-nowrap shrink-0 cursor-pointer"
                  >
                    Xóa toàn bộ lịch sử
                  </button>
                )}
              </div>

              {exportHistory.length > 0 && (
                <div className="px-6 py-3.5 bg-slate-50/80 border-b border-slate-200 flex flex-col lg:flex-row lg:items-center justify-between gap-3">
                  <div className="relative flex-1 max-w-md">
                    <Search className="w-4 h-4 text-slate-400 absolute left-3 top-1/2 -translate-y-1/2" />
                    <input
                      type="text"
                      value={historySearchQuery}
                      onChange={(e) => setHistorySearchQuery(e.target.value)}
                      placeholder="Tìm theo tên file, tên mẫu hoặc nội dung đã điền (VD: tên người, số tiền)..."
                      className="w-full pl-9 pr-8 py-1.5 text-xs bg-white border border-slate-300 rounded-lg text-slate-900 focus:outline-none focus:border-blue-600"
                    />
                    {historySearchQuery && (
                      <button
                        type="button"
                        onClick={() => setHistorySearchQuery('')}
                        className="absolute right-2.5 top-1/2 -translate-y-1/2 text-slate-400 hover:text-slate-700 cursor-pointer"
                      >
                        <X className="w-3.5 h-3.5" />
                      </button>
                    )}
                  </div>

                  <div className="flex flex-wrap items-center gap-2.5">
                    <div className="flex items-center gap-1.5">
                      <Filter className="w-3.5 h-3.5 text-slate-500 shrink-0" />
                      <select
                        value={historyTemplateFilter}
                        onChange={(e) => setHistoryTemplateFilter(e.target.value)}
                        className="px-2.5 py-1.5 text-xs bg-white border border-slate-300 rounded-lg text-slate-700 focus:outline-none focus:border-blue-600"
                      >
                        <option value="ALL">Tất cả biểu mẫu gốc</option>
                        {historyTemplateOptions.map((tplName) => (
                          <option key={tplName} value={tplName}>
                            Mẫu: {tplName}
                          </option>
                        ))}
                      </select>
                    </div>

                    <div className="flex items-center gap-1 p-1 bg-slate-200/70 rounded-lg">
                      <button
                        type="button"
                        onClick={() => setHistoryTimeFilter('all')}
                        className={`px-2.5 py-1 text-xs font-medium rounded-md transition-colors whitespace-nowrap shrink-0 cursor-pointer ${
                          historyTimeFilter === 'all'
                            ? 'bg-white text-slate-900 shadow-2xs font-semibold'
                            : 'text-slate-600 hover:text-slate-900'
                        }`}
                      >
                        Mọi lúc
                      </button>
                      <button
                        type="button"
                        onClick={() => setHistoryTimeFilter('today')}
                        className={`px-2.5 py-1 text-xs font-medium rounded-md transition-colors whitespace-nowrap shrink-0 cursor-pointer ${
                          historyTimeFilter === 'today'
                            ? 'bg-white text-slate-900 shadow-2xs font-semibold'
                            : 'text-slate-600 hover:text-slate-900'
                        }`}
                      >
                        Hôm nay
                      </button>
                      <button
                        type="button"
                        onClick={() => setHistoryTimeFilter('7days')}
                        className={`px-2.5 py-1 text-xs font-medium rounded-md transition-colors whitespace-nowrap shrink-0 cursor-pointer ${
                          historyTimeFilter === '7days'
                            ? 'bg-white text-slate-900 shadow-2xs font-semibold'
                            : 'text-slate-600 hover:text-slate-900'
                        }`}
                      >
                        7 ngày qua
                      </button>
                      <button
                        type="button"
                        onClick={() => setHistoryTimeFilter('30days')}
                        className={`px-2.5 py-1 text-xs font-medium rounded-md transition-colors whitespace-nowrap shrink-0 cursor-pointer ${
                          historyTimeFilter === '30days'
                            ? 'bg-white text-slate-900 shadow-2xs font-semibold'
                            : 'text-slate-600 hover:text-slate-900'
                        }`}
                      >
                        30 ngày qua
                      </button>
                    </div>

                    {(historySearchQuery ||
                      historyTemplateFilter !== 'ALL' ||
                      historyTimeFilter !== 'all') && (
                      <button
                        type="button"
                        onClick={() => {
                          setHistorySearchQuery('');
                          setHistoryTemplateFilter('ALL');
                          setHistoryTimeFilter('all');
                        }}
                        className="text-xs font-medium text-blue-700 hover:text-blue-800 whitespace-nowrap shrink-0 cursor-pointer"
                      >
                        Xóa bộ lọc
                      </button>
                    )}
                  </div>
                </div>
              )}

              {exportHistory.length === 0 ? (
                <div className="p-8 text-center text-xs text-slate-500">
                  Chưa có file nào được xuất. Khi bạn nhấn &ldquo;Xuất File Word Hoàn Chỉnh&rdquo;, bản ghi sẽ lưu tại đây để tìm kiếm và tải lại nhanh.
                </div>
              ) : filteredExportHistory.length === 0 ? (
                <div className="p-8 text-center space-y-2">
                  <p className="text-sm font-medium text-slate-700">
                    Không tìm thấy lượt xuất file nào khớp với điều kiện lọc.
                  </p>
                  <button
                    type="button"
                    onClick={() => {
                      setHistorySearchQuery('');
                      setHistoryTemplateFilter('ALL');
                      setHistoryTimeFilter('all');
                    }}
                    className="px-3 py-1.5 text-xs font-semibold text-blue-700 bg-blue-50 hover:bg-blue-100 rounded-md cursor-pointer"
                  >
                    Đặt lại bộ lọc lịch sử
                  </button>
                </div>
              ) : (
                <div className="overflow-x-auto">
                  <table className="w-full text-left border-collapse">
                    <thead>
                      <tr className="border-b border-slate-200 bg-slate-50 text-xs font-semibold text-slate-600">
                        <th className="py-3 px-5">Tên file xuất</th>
                        <th className="py-3 px-4">Biểu mẫu gốc</th>
                        <th className="py-3 px-4">Tóm tắt dữ liệu đã điền</th>
                        <th className="py-3 px-4 text-right">Số dòng đổi</th>
                        <th className="py-3 px-4 text-right">Thời gian</th>
                        <th className="py-3 px-5 text-right">Thao tác</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-200 text-sm">
                      {filteredExportHistory.map((rec) => {
                        const previewValues = rec.fieldValues
                          ? Object.entries(rec.fieldValues)
                              .filter(([, v]) => v && v.trim().length > 0)
                              .slice(0, 3)
                              .map(([k, v]) => `${k}: ${v}`)
                              .join(' · ')
                          : '';

                        return (
                          <tr key={rec.id} className="hover:bg-slate-50/80 transition-colors">
                            <td className="py-3 px-5 font-mono-tabular text-xs font-medium text-slate-900">
                              {rec.outputFileName}
                            </td>
                            <td className="py-3 px-4 text-xs text-slate-700">{rec.templateName}</td>
                            <td className="py-3 px-4 text-xs text-slate-500 max-w-xs truncate">
                              {previewValues || 'Đã điền theo dòng văn bản'}
                            </td>
                            <td className="py-3 px-4 text-right font-mono-tabular text-xs text-emerald-700 font-medium">
                              {rec.modifiedCount} dòng
                            </td>
                            <td className="py-3 px-4 text-right font-mono-tabular text-xs text-slate-500">
                              {new Date(rec.exportedAt).toLocaleString('vi-VN')}
                            </td>
                            <td className="py-3 px-5 text-right">
                              <div className="inline-flex items-center gap-2">
                                {templatesConfig[rec.templateName] && (
                                  <>
                                    <button
                                      type="button"
                                      onClick={() => {
                                        setChosenTemplate(rec.templateName);
                                        setActiveSection('use');
                                        setTimeout(() => {
                                          if (rec.fieldValues) {
                                            setMailingsFieldValues(rec.fieldValues);
                                          }
                                          setUserReplacements(rec.replacements);
                                        }, 80);
                                      }}
                                      className="px-2.5 py-1 text-xs font-medium text-blue-700 bg-blue-50 hover:bg-blue-100 rounded whitespace-nowrap shrink-0 cursor-pointer"
                                    >
                                      Nạp lại dữ liệu
                                    </button>
                                    <button
                                      type="button"
                                      onClick={() =>
                                        handleQuickUploadDocxToDrive(
                                          rec.templateName,
                                          rec.replacements,
                                          rec.outputFileName
                                        )
                                      }
                                      title="Lưu file này lên Google Drive"
                                      className="inline-flex items-center gap-1 px-2.5 py-1 text-xs font-medium text-slate-700 bg-slate-100 hover:bg-slate-200 rounded whitespace-nowrap shrink-0 cursor-pointer"
                                    >
                                      <CloudUpload className="w-3.5 h-3.5 text-blue-700" />
                                      <span>Lưu Drive</span>
                                    </button>
                                    <button
                                      type="button"
                                      onClick={async () => {
                                        const info = templatesConfig[rec.templateName];
                                        if (!info) return;
                                        const blob = await exportFilledDocxBlob(
                                          info.base64_docx,
                                          rec.replacements
                                        );
                                        triggerBlobDownload(blob, rec.outputFileName);
                                      }}
                                      className="px-2.5 py-1 text-xs font-medium text-slate-700 bg-slate-100 hover:bg-slate-200 rounded whitespace-nowrap shrink-0 cursor-pointer"
                                    >
                                      Tải .docx
                                    </button>
                                  </>
                                )}
                              </div>
                            </td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>
              )}
            </section>
          </div>
        )}

        {/* ================= TAB 4: CẤU HÌNH CONFIG.JSON ================= */}
        {activeSection === 'json' && (
          <div className="grid grid-cols-1 lg:grid-cols-12 gap-8 items-start">
            <div className="lg:col-span-7 bg-white border border-slate-200 rounded-lg p-6 space-y-4">
              <div className="flex flex-wrap items-center justify-between gap-2">
                <div>
                  <h2 className="text-base font-semibold text-slate-900">
                    Cấu trúc tệp templates_store/config.json
                  </h2>
                  <p className="text-xs text-slate-500 mt-0.5">
                    Lưu trữ danh sách dòng (selected_indices) và mẫu vị trí trường Mailings (paragraph_templates)
                  </p>
                </div>

                <div className="flex items-center gap-2">
                  <button
                    type="button"
                    onClick={() => {
                      navigator.clipboard.writeText(formatPythonCompatibleConfig(templatesConfig));
                      setCopiedJson(true);
                      setTimeout(() => setCopiedJson(false), 2000);
                    }}
                    className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-medium text-slate-700 bg-slate-100 hover:bg-slate-200 rounded-md whitespace-nowrap shrink-0 cursor-pointer"
                  >
                    {copiedJson ? (
                      <>
                        <Check className="w-3.5 h-3.5 text-emerald-600" />
                        <span>Đã sao chép</span>
                      </>
                    ) : (
                      <>
                        <Copy className="w-3.5 h-3.5" />
                        <span>Sao chép config.json</span>
                      </>
                    )}
                  </button>

                  <button
                    type="button"
                    onClick={() => {
                      const jsonBlob = new Blob([JSON.stringify(templatesConfig, null, 2)], {
                        type: 'application/json',
                      });
                      triggerBlobDownload(jsonBlob, 'templates_store_backup.json');
                    }}
                    className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold text-white bg-blue-700 hover:bg-blue-800 rounded-md whitespace-nowrap shrink-0 cursor-pointer"
                  >
                    <Download className="w-3.5 h-3.5" />
                    <span>Tải bản sao lưu đầy đủ</span>
                  </button>
                </div>
              </div>

              <pre className="p-4 bg-slate-900 text-slate-100 rounded-lg text-xs font-mono-tabular overflow-x-auto max-h-[500px] leading-relaxed">
                {formatPythonCompatibleConfig(templatesConfig)}
              </pre>
            </div>

            <div className="lg:col-span-5 bg-white border border-slate-200 rounded-lg p-6 space-y-4">
              <h3 className="text-base font-semibold text-slate-900">
                Khôi phục hoặc đồng bộ thư viện
              </h3>
              <p className="text-xs text-slate-600 leading-relaxed">
                Bạn có thể dán nội dung JSON sao lưu vào ô bên dưới hoặc đồng bộ trực tiếp với Google Drive.
              </p>

              <textarea
                rows={7}
                placeholder='Dán nội dung JSON hoặc {"Tên mẫu": {"selected_indices": ["1", "3"]}} vào đây...'
                value={jsonImportText}
                onChange={(e) => setJsonImportText(e.target.value)}
                className="w-full p-3 text-xs font-mono-tabular bg-slate-50 border border-slate-300 rounded-lg text-slate-900 focus:outline-none focus:bg-white focus:border-blue-600"
              />

              {jsonStatusMessage && (
                <p className="text-xs font-medium text-blue-800 bg-blue-50 border border-blue-200 rounded p-2.5">
                  {jsonStatusMessage}
                </p>
              )}

              <div className="flex flex-wrap items-center gap-3 pt-1">
                <button
                  type="button"
                  onClick={() => {
                    try {
                      const parsed = JSON.parse(jsonImportText);
                      if (typeof parsed !== 'object' || parsed === null) {
                        throw new Error('JSON không hợp lệ.');
                      }
                      const merged = { ...templatesConfig };
                      Object.entries(parsed).forEach(([k, val]: [string, any]) => {
                        if (val && Array.isArray(val.selected_indices)) {
                          if (val.base64_docx) {
                            merged[k] = val;
                          } else if (merged[k]) {
                            merged[k] = {
                              ...merged[k],
                              selected_indices: val.selected_indices.map(String),
                              ...(val.paragraph_templates
                                ? { paragraph_templates: val.paragraph_templates }
                                : {}),
                            };
                          }
                        }
                      });
                      saveTemplatesConfig(merged);
                      setTemplatesConfig(merged);
                      setJsonStatusMessage('Đã cập nhật cấu hình JSON thành công!');
                    } catch {
                      setJsonStatusMessage('Lỗi: Chuỗi JSON không đúng định dạng.');
                    }
                  }}
                  className="px-4 py-2 text-xs font-semibold text-white bg-blue-700 hover:bg-blue-800 rounded-lg whitespace-nowrap shrink-0 cursor-pointer"
                >
                  Nhập cấu hình JSON
                </button>

                <button
                  type="button"
                  onClick={() => setIsDriveModalOpen(true)}
                  className="inline-flex items-center gap-1.5 px-3.5 py-2 text-xs font-semibold text-blue-700 bg-blue-50 hover:bg-blue-100 rounded-lg whitespace-nowrap shrink-0 cursor-pointer"
                >
                  <Cloud className="w-3.5 h-3.5" />
                  <span>Đồng bộ qua Google Drive</span>
                </button>

                <button
                  type="button"
                  onClick={async () => {
                    localStorage.removeItem('smart_word_templates_seeded_v2');
                    localStorage.removeItem('smart_word_templates_config_v2');
                    const fresh = await ensureStarterTemplatesSeeded();
                    setTemplatesConfig(fresh);
                    setChosenTemplate(Object.keys(fresh)[0] || '');
                    setJsonStatusMessage('Đã khôi phục 3 biểu mẫu hành chính mặc định.');
                  }}
                  className="inline-flex items-center gap-1.5 px-3.5 py-2 text-xs font-medium text-slate-700 bg-slate-100 hover:bg-slate-200 rounded-lg whitespace-nowrap shrink-0 cursor-pointer"
                >
                  <RefreshCw className="w-3.5 h-3.5" />
                  <span>Khôi phục 3 mẫu mặc định</span>
                </button>
              </div>
            </div>
          </div>
        )}
      </main>

      {/* Quiet Footer */}
      <footer className="mt-auto border-t border-slate-200 bg-white py-4 px-6 lg:px-10">
        <div className="max-w-[1440px] mx-auto flex flex-col sm:flex-row items-center justify-between gap-2 text-xs text-slate-500">
          <span>Thư Viện Biểu Mẫu Word Thông Minh · Tích hợp Google Drive & Trộn Thư Mailings («Trường»)</span>
          <div className="flex items-center gap-3">
            <button
              type="button"
              onClick={() => setActiveSection('library')}
              className="hover:text-slate-900 cursor-pointer"
            >
              Kho biểu mẫu
            </button>
            <span aria-hidden="true">·</span>
            <button
              type="button"
              onClick={() => setActiveSection('use')}
              className="hover:text-slate-900 cursor-pointer"
            >
              Sử dụng mẫu
            </button>
            <span aria-hidden="true">·</span>
            <button
              type="button"
              onClick={() => setActiveSection('configure')}
              className="hover:text-slate-900 cursor-pointer"
            >
              Thêm & Cấu hình vị trí trường
            </button>
            <span aria-hidden="true">·</span>
            <button
              type="button"
              onClick={() => setIsDriveModalOpen(true)}
              className="hover:text-slate-900 cursor-pointer"
            >
              Google Drive
            </button>
          </div>
        </div>
      </footer>
    </div>
  );
}
