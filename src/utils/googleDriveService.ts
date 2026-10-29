import { initializeApp } from 'firebase/app';
import {
  getAuth,
  signInWithPopup,
  GoogleAuthProvider,
  onAuthStateChanged,
  User,
} from 'firebase/auth';
import firebaseConfig from '../../firebase-applet-config.json';
import { exportFilledDocxBlob } from './docxEngine';
import {
  SharedDrivePublicationInfo,
  TemplatesConfigMap,
} from './templateStore';

export const SCOPES = [
  'https://www.googleapis.com/auth/drive',
  'https://www.googleapis.com/auth/drive.activity',
  'https://www.googleapis.com/auth/drive.activity.readonly',
  'https://www.googleapis.com/auth/drive.appdata',
  'https://www.googleapis.com/auth/drive.apps.readonly',
  'https://www.googleapis.com/auth/drive.file',
  'https://www.googleapis.com/auth/drive.install',
  'https://www.googleapis.com/auth/drive.meet.readonly',
  'https://www.googleapis.com/auth/drive.metadata',
  'https://www.googleapis.com/auth/drive.metadata.readonly',
  'https://www.googleapis.com/auth/drive.photos.readonly',
  'https://www.googleapis.com/auth/drive.readonly',
  'https://www.googleapis.com/auth/drive.scripts',
];

const app = initializeApp(firebaseConfig);
const auth = getAuth(app);

const provider = new GoogleAuthProvider();
SCOPES.forEach((scope) => {
  provider.addScope(scope);
});

// Flag to indicate if we are in the middle of a sign-in flow.
let isSigningIn = false;
// Cache the access token strictly in memory (never in localStorage/sessionStorage).
let cachedAccessToken: string | null = null;

export const initAuth = (
  onAuthSuccess?: (user: User, token: string) => void,
  onAuthFailure?: () => void
) => {
  return onAuthStateChanged(auth, async (user: User | null) => {
    if (user) {
      if (cachedAccessToken) {
        if (onAuthSuccess) onAuthSuccess(user, cachedAccessToken);
      } else if (!isSigningIn) {
        cachedAccessToken = null;
        if (onAuthFailure) onAuthFailure();
      }
    } else {
      cachedAccessToken = null;
      if (onAuthFailure) onAuthFailure();
    }
  });
};

export const googleSignIn = async (): Promise<{
  user: User;
  accessToken: string;
} | null> => {
  try {
    isSigningIn = true;
    const result = await signInWithPopup(auth, provider);
    const credential = GoogleAuthProvider.credentialFromResult(result);
    if (!credential?.accessToken) {
      throw new Error('Không lấy được OAuth access token từ Google.');
    }

    cachedAccessToken = credential.accessToken;
    return { user: result.user, accessToken: cachedAccessToken };
  } catch (error: any) {
    console.error('Sign in error:', error);
    throw error;
  } finally {
    isSigningIn = false;
  }
};

export const getAccessToken = async (): Promise<string | null> => {
  return cachedAccessToken;
};

export const logoutGoogle = async () => {
  await auth.signOut();
  cachedAccessToken = null;
};

export interface DriveFileItem {
  id: string;
  name: string;
  mimeType: string;
  modifiedTime?: string;
  size?: string;
  webViewLink?: string;
}

const DOCX_MIME =
  'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
const GOOGLE_DOC_MIME = 'application/vnd.google-apps.document';
const FOLDER_MIME = 'application/vnd.google-apps.folder';
export const SHARED_DRIVE_FOLDER_NAME = 'Thư Viện Biểu Mẫu Word (Dùng Chung App)';
export const SHARED_BACKUP_FILE_NAME = 'templates_store_backup.json';

/**
 * Extracts a Google Drive File ID from either a full Drive Share URL or a raw File ID string.
 */
export function extractDriveFileIdFromInput(input: string): string | null {
  const trimmed = input.trim();
  if (!trimmed) return null;

  // Match /file/d/{id} or /folders/{id}
  const pathMatch = trimmed.match(/\/(?:file\/d|folders|d)\/([a-zA-Z0-9_-]{10,})/);
  if (pathMatch && pathMatch[1]) {
    return pathMatch[1];
  }

  // Match ?id={id} or &id={id} or ?driveId={id}
  const queryMatch = trimmed.match(/[?&](?:id|driveId)=([a-zA-Z0-9_-]{10,})/);
  if (queryMatch && queryMatch[1]) {
    return queryMatch[1];
  }

  // Raw file ID
  if (/^[a-zA-Z0-9_-]{10,}$/.test(trimmed)) {
    return trimmed;
  }

  return null;
}

/**
 * Lists Word (.docx), Google Docs, and JSON backup files from the user's Google Drive.
 */
export async function listDriveWordAndBackupFiles(
  searchQuery = ''
): Promise<DriveFileItem[]> {
  const token = await getAccessToken();
  if (!token) {
    throw new Error('UNAUTHENTICATED');
  }

  const mimeFilter = `(mimeType='${DOCX_MIME}' or mimeType='${GOOGLE_DOC_MIME}' or name contains 'templates_store_backup')`;
  const escapedSearch = searchQuery.trim().replace(/'/g, "\\'");
  const nameFilter = escapedSearch ? ` and name contains '${escapedSearch}'` : '';
  const q = `trashed = false and ${mimeFilter}${nameFilter}`;

  const params = new URLSearchParams({
    q,
    pageSize: '30',
    orderBy: 'modifiedTime desc',
    fields: 'files(id,name,mimeType,modifiedTime,size,webViewLink)',
  });

  const res = await fetch(
    `https://www.googleapis.com/drive/v3/files?${params.toString()}`,
    {
      headers: {
        Authorization: `Bearer ${token}`,
      },
    }
  );

  if (res.status === 401 || res.status === 403) {
    throw new Error('UNAUTHENTICATED');
  }

  if (!res.ok) {
    const errData = await res.json().catch(() => ({}));
    throw new Error(
      errData?.error?.message || 'Không thể tải danh sách tệp từ Google Drive.'
    );
  }

  const data = await res.json();
  return (data.files || []) as DriveFileItem[];
}

/**
 * Downloads a .docx file (or exports a Google Doc as .docx) from Google Drive as an ArrayBuffer.
 */
export async function downloadDriveDocxAsArrayBuffer(
  file: DriveFileItem
): Promise<{ buffer: ArrayBuffer; fileName: string }> {
  const token = await getAccessToken();
  if (!token) {
    throw new Error('UNAUTHENTICATED');
  }

  const isGoogleDoc = file.mimeType === GOOGLE_DOC_MIME;
  const endpoint = isGoogleDoc
    ? `https://www.googleapis.com/drive/v3/files/${encodeURIComponent(
        file.id
      )}/export?mimeType=${encodeURIComponent(DOCX_MIME)}`
    : `https://www.googleapis.com/drive/v3/files/${encodeURIComponent(
        file.id
      )}?alt=media`;

  const res = await fetch(endpoint, {
    headers: {
      Authorization: `Bearer ${token}`,
    },
  });

  if (res.status === 401 || res.status === 403) {
    throw new Error('UNAUTHENTICATED');
  }

  if (!res.ok) {
    throw new Error(`Không thể tải tệp "${file.name}" từ Google Drive.`);
  }

  const buffer = await res.arrayBuffer();
  const cleanName = file.name.toLowerCase().endsWith('.docx')
    ? file.name
    : `${file.name}.docx`;

  return { buffer, fileName: cleanName };
}

/**
 * Downloads a JSON backup file from Google Drive using the authenticated user's token.
 */
export async function downloadDriveJsonFile(fileId: string): Promise<string> {
  const token = await getAccessToken();
  if (!token) {
    throw new Error('UNAUTHENTICATED');
  }

  const res = await fetch(
    `https://www.googleapis.com/drive/v3/files/${encodeURIComponent(
      fileId
    )}?alt=media`,
    {
      headers: {
        Authorization: `Bearer ${token}`,
      },
    }
  );

  if (!res.ok) {
    throw new Error('Không thể đọc tệp JSON từ Google Drive.');
  }

  return await res.text();
}

/**
 * Uploads a new file or updates (overwrites) an existing file on Google Drive.
 * Note: Callers MUST display a confirmation dialog before passing `existingFileId` to overwrite user data!
 */
export async function uploadOrUpdateFileToDrive(params: {
  fileName: string;
  blob: Blob;
  mimeType: string;
  existingFileId?: string;
  parentFolderId?: string;
}): Promise<DriveFileItem> {
  const token = await getAccessToken();
  if (!token) {
    throw new Error('UNAUTHENTICATED');
  }

  const metadata: Record<string, any> = {
    name: params.fileName,
    mimeType: params.mimeType,
  };

  if (!params.existingFileId && params.parentFolderId) {
    metadata.parents = [params.parentFolderId];
  }

  const form = new FormData();
  form.append(
    'metadata',
    new Blob([JSON.stringify(metadata)], { type: 'application/json' })
  );
  form.append('file', params.blob);

  const url = params.existingFileId
    ? `https://www.googleapis.com/upload/drive/v3/files/${encodeURIComponent(
        params.existingFileId
      )}?uploadType=multipart&fields=id,name,mimeType,modifiedTime,size,webViewLink`
    : 'https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&fields=id,name,mimeType,modifiedTime,size,webViewLink';

  const method = params.existingFileId ? 'PATCH' : 'POST';

  const res = await fetch(url, {
    method,
    headers: {
      Authorization: `Bearer ${token}`,
    },
    body: form,
  });

  if (res.status === 401 || res.status === 403) {
    throw new Error('UNAUTHENTICATED');
  }

  if (!res.ok) {
    const errData = await res.json().catch(() => ({}));
    throw new Error(
      errData?.error?.message || 'Lỗi khi tải tệp lên Google Drive.'
    );
  }

  return (await res.json()) as DriveFileItem;
}

/**
 * Grants "Anyone with the link can view" (`role: 'reader', type: 'anyone'`) permission
 * on a Google Drive file or folder so that any user accessing the app can read the shared templates.
 * Note: Callers MUST present a confirmation dialog before invoking this on user Drive data.
 */
export async function shareDriveFileWithAnyone(fileId: string): Promise<void> {
  const token = await getAccessToken();
  if (!token) {
    throw new Error('UNAUTHENTICATED');
  }

  const res = await fetch(
    `https://www.googleapis.com/drive/v3/files/${encodeURIComponent(
      fileId
    )}/permissions`,
    {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({
        role: 'reader',
        type: 'anyone',
      }),
    }
  );

  if (res.status === 401 || res.status === 403) {
    throw new Error('UNAUTHENTICATED');
  }
}

/**
 * Finds or creates the dedicated shared folder on Google Drive (`Thư Viện Biểu Mẫu Word (Dùng Chung App)`).
 */
export async function findOrCreateSharedDriveFolder(
  folderName = SHARED_DRIVE_FOLDER_NAME
): Promise<DriveFileItem> {
  const token = await getAccessToken();
  if (!token) {
    throw new Error('UNAUTHENTICATED');
  }

  const escapedName = folderName.replace(/'/g, "\\'");
  const q = `trashed = false and mimeType = '${FOLDER_MIME}' and name = '${escapedName}'`;
  const searchParams = new URLSearchParams({
    q,
    pageSize: '5',
    fields: 'files(id,name,mimeType,webViewLink)',
  });

  const searchRes = await fetch(
    `https://www.googleapis.com/drive/v3/files?${searchParams.toString()}`,
    {
      headers: {
        Authorization: `Bearer ${token}`,
      },
    }
  );

  if (searchRes.status === 401 || searchRes.status === 403) {
    throw new Error('UNAUTHENTICATED');
  }

  if (searchRes.ok) {
    const data = await searchRes.json();
    if (data.files && data.files.length > 0) {
      return data.files[0] as DriveFileItem;
    }
  }

  // Create new folder on Google Drive
  const createRes = await fetch(
    'https://www.googleapis.com/drive/v3/files?fields=id,name,mimeType,webViewLink',
    {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({
        name: folderName,
        mimeType: FOLDER_MIME,
      }),
    }
  );

  if (!createRes.ok) {
    throw new Error('Không thể tạo thư mục dùng chung trên Google Drive.');
  }

  return (await createRes.json()) as DriveFileItem;
}

/**
 * Publishes the entire template library (`templates_store_backup.json` + individual `.docx` files)
 * to a shared Google Drive folder, sets `type: 'anyone', role: 'reader'` permissions so anyone
 * visiting the app can access it, and registers the Drive File ID with the app backend.
 *
 * MANDATORY: Must only be called after the user confirms the operation in the confirmation dialog!
 */
export async function publishLibraryToSharedGoogleDrive(params: {
  templatesConfig: TemplatesConfigMap;
  ownerEmail?: string;
  existingBackupFileId?: string;
  uploadIndividualDocxFiles?: boolean;
}): Promise<{
  driveInfo: SharedDrivePublicationInfo;
  updatedTemplatesConfig: TemplatesConfigMap;
}> {
  const token = await getAccessToken();
  if (!token) {
    throw new Error('UNAUTHENTICATED');
  }

  // 1. Find or create the shared folder on Google Drive and share it with anyone (reader)
  const folder = await findOrCreateSharedDriveFolder(SHARED_DRIVE_FOLDER_NAME);
  await shareDriveFileWithAnyone(folder.id);

  // 2. Optionally upload individual .docx files for each template into the shared Drive folder
  const updatedTemplatesConfig: TemplatesConfigMap = { ...params.templatesConfig };

  if (params.uploadIndividualDocxFiles !== false) {
    for (const [tplName, entry] of Object.entries(updatedTemplatesConfig)) {
      try {
        const docxBlob = await exportFilledDocxBlob(entry.base64_docx, {});
        const rawFileName =
          entry.file_path.split('/').pop() || `${tplName}.docx`;

        const uploadedDocx = await uploadOrUpdateFileToDrive({
          fileName: rawFileName,
          blob: docxBlob,
          mimeType: DOCX_MIME,
          existingFileId: entry.drive_file_id,
          parentFolderId: folder.id,
        });

        await shareDriveFileWithAnyone(uploadedDocx.id);

        updatedTemplatesConfig[tplName] = {
          ...entry,
          drive_file_id: uploadedDocx.id,
          drive_web_view_link:
            uploadedDocx.webViewLink ||
            `https://drive.google.com/file/d/${uploadedDocx.id}/view?usp=sharing`,
        };
      } catch (err) {
        console.warn(`Could not upload individual docx for "${tplName}":`, err);
      }
    }
  }

  // 3. Look up if `templates_store_backup.json` already exists in Drive if existingBackupFileId wasn't passed
  let targetBackupFileId = params.existingBackupFileId;
  if (!targetBackupFileId) {
    try {
      const existingFiles = await listDriveWordAndBackupFiles(
        SHARED_BACKUP_FILE_NAME
      );
      const found = existingFiles.find(
        (f) => f.name === SHARED_BACKUP_FILE_NAME
      );
      if (found) {
        targetBackupFileId = found.id;
      }
    } catch {
      // Ignore
    }
  }

  // 4. Upload or update `templates_store_backup.json` inside the shared Drive folder
  const jsonBlob = new Blob(
    [JSON.stringify(updatedTemplatesConfig, null, 2)],
    {
      type: 'application/json',
    }
  );

  const uploadedBackup = await uploadOrUpdateFileToDrive({
    fileName: SHARED_BACKUP_FILE_NAME,
    blob: jsonBlob,
    mimeType: 'application/json',
    existingFileId: targetBackupFileId,
    parentFolderId: folder.id,
  });

  // 5. Grant public reader access on `templates_store_backup.json` so any visitor to the app can read it from Drive
  await shareDriveFileWithAnyone(uploadedBackup.id);

  const driveInfo: SharedDrivePublicationInfo = {
    backupFileId: uploadedBackup.id,
    backupFileName: uploadedBackup.name,
    backupWebViewLink:
      uploadedBackup.webViewLink ||
      `https://drive.google.com/file/d/${uploadedBackup.id}/view?usp=sharing`,
    folderId: folder.id,
    folderName: folder.name,
    folderWebViewLink:
      folder.webViewLink ||
      `https://drive.google.com/drive/folders/${folder.id}?usp=sharing`,
    ownerEmail: params.ownerEmail,
    updatedAt: new Date().toISOString(),
    templateCount: Object.keys(updatedTemplatesConfig).length,
  };

  // 6. Register the shared Google Drive file ID and templates with the server so all app visitors load from Drive
  await fetch('/api/drive-config', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      driveConfig: driveInfo,
      templates: updatedTemplatesConfig,
    }),
  });

  return {
    driveInfo,
    updatedTemplatesConfig,
  };
}

/**
 * Deletes a file from Google Drive.
 * MANDATORY: Callers MUST ask for explicit user confirmation via modal before calling this function!
 */
export async function deleteFileFromDrive(fileId: string): Promise<void> {
  const token = await getAccessToken();
  if (!token) {
    throw new Error('UNAUTHENTICATED');
  }

  const res = await fetch(
    `https://www.googleapis.com/drive/v3/files/${encodeURIComponent(fileId)}`,
    {
      method: 'DELETE',
      headers: {
        Authorization: `Bearer ${token}`,
      },
    }
  );

  if (res.status === 401 || res.status === 403) {
    throw new Error('UNAUTHENTICATED');
  }

  if (!res.ok) {
    throw new Error('Không thể xóa tệp khỏi Google Drive.');
  }
}
