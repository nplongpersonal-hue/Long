import { generateSampleDocxBase64, SampleParagraphSpec } from './docxEngine';

export type MailingsFieldCategory =
  | 'Cá nhân & Nhân sự'
  | 'Đơn vị & Tổ chức'
  | 'Thời gian & Địa điểm'
  | 'Tài chính & Văn bản'
  | 'Trường tùy chỉnh';

export interface MailingsFieldDef {
  key: string;
  label: string;
  category: MailingsFieldCategory;
  placeholder: string;
  defaultValue?: string;
}

export const DEFAULT_MAILINGS_FIELDS: MailingsFieldDef[] = [
  // 1. Cá nhân & Nhân sự
  {
    key: 'Ho_Va_Ten',
    label: 'Họ và tên',
    category: 'Cá nhân & Nhân sự',
    placeholder: 'Nguyễn Văn Minh',
  },
  {
    key: 'Danh_Xung',
    label: 'Danh xưng (Ông/Bà)',
    category: 'Cá nhân & Nhân sự',
    placeholder: 'Ông',
  },
  {
    key: 'Ma_Nhan_Vien',
    label: 'Mã nhân viên',
    category: 'Cá nhân & Nhân sự',
    placeholder: 'NV-2048',
  },
  {
    key: 'Chuc_Vu',
    label: 'Chức vụ',
    category: 'Cá nhân & Nhân sự',
    placeholder: 'Chuyên viên Phân tích Nghiệp vụ',
  },
  {
    key: 'Phong_Ban',
    label: 'Phòng ban / Bộ phận',
    category: 'Cá nhân & Nhân sự',
    placeholder: 'Phòng Triển khai Dự án Công nghệ',
  },
  {
    key: 'So_Dien_Thoai',
    label: 'Số điện thoại',
    category: 'Cá nhân & Nhân sự',
    placeholder: '0912.345.678',
  },
  {
    key: 'Email',
    label: 'Địa chỉ Email',
    category: 'Cá nhân & Nhân sự',
    placeholder: 'minh.nv@congty.vn',
  },
  {
    key: 'CCCD_CMND',
    label: 'Số CCCD / CMND',
    category: 'Cá nhân & Nhân sự',
    placeholder: '001095012345',
  },
  {
    key: 'Dia_Chi',
    label: 'Địa chỉ liên hệ',
    category: 'Cá nhân & Nhân sự',
    placeholder: 'Số 15 Duy Tân, Cầu Giấy, Hà Nội',
  },

  // 2. Đơn vị & Tổ chức
  {
    key: 'Kinh_Gui',
    label: 'Kính gửi (Đơn vị/Cấp trên)',
    category: 'Đơn vị & Tổ chức',
    placeholder: 'Ban Giám đốc và Phòng Tài chính - Kế toán',
  },
  {
    key: 'Ten_Cong_Ty',
    label: 'Tên Công ty / Đơn vị',
    category: 'Đơn vị & Tổ chức',
    placeholder: 'CÔNG TY CỔ PHẦN GIẢI PHÁP SỐ ĐÔNG DƯƠNG',
  },
  {
    key: 'Ben_Giao',
    label: 'Bên giao / Người giao',
    category: 'Đơn vị & Tổ chức',
    placeholder: 'Ông Phạm Quốc Khánh - Chuyên viên Quản trị Hệ thống',
  },
  {
    key: 'Ben_Nhan',
    label: 'Bên nhận / Người nhận bàn giao',
    category: 'Đơn vị & Tổ chức',
    placeholder: 'Bà Vũ Ngọc Mai - Phòng Kinh doanh Doanh nghiệp',
  },

  // 3. Thời gian & Địa điểm
  {
    key: 'Dia_Diem',
    label: 'Địa điểm lập văn bản',
    category: 'Thời gian & Địa điểm',
    placeholder: 'Hà Nội',
  },
  {
    key: 'Ngay',
    label: 'Ngày',
    category: 'Thời gian & Địa điểm',
    placeholder: '23',
  },
  {
    key: 'Thang',
    label: 'Tháng',
    category: 'Thời gian & Địa điểm',
    placeholder: '10',
  },
  {
    key: 'Nam',
    label: 'Năm',
    category: 'Thời gian & Địa điểm',
    placeholder: '2026',
  },
  {
    key: 'Thoi_Gian_Tu_Den',
    label: 'Khoảng thời gian (Từ ngày... đến ngày...)',
    category: 'Thời gian & Địa điểm',
    placeholder: 'Từ ngày 26/10/2026 đến hết ngày 28/10/2026 (03 ngày)',
  },

  // 4. Tài chính & Văn bản
  {
    key: 'So_Hieu_Van_Ban',
    label: 'Số hiệu văn bản',
    category: 'Tài chính & Văn bản',
    placeholder: '42/2026/BBBG-HCNS',
  },
  {
    key: 'Noi_Dung',
    label: 'Nội dung công việc / Thanh toán',
    category: 'Tài chính & Văn bản',
    placeholder: 'Chi phí mua bản quyền máy chủ và tên miền quý IV/2026',
  },
  {
    key: 'Ly_Do',
    label: 'Lý do đề nghị / Nghỉ phép',
    category: 'Tài chính & Văn bản',
    placeholder: 'Giải quyết công việc gia đình tại quê nhà.',
  },
  {
    key: 'So_Tien_So',
    label: 'Số tiền bằng số',
    category: 'Tài chính & Văn bản',
    placeholder: '18.500.000 VNĐ',
  },
  {
    key: 'So_Tien_Chu',
    label: 'Số tiền bằng chữ',
    category: 'Tài chính & Văn bản',
    placeholder: 'Mười tám triệu năm trăm nghìn đồng chẵn.',
  },
  {
    key: 'Hinh_Thuc_Thanh_Toan',
    label: 'Hình thức & Số tài khoản',
    category: 'Tài chính & Văn bản',
    placeholder: 'Chuyển khoản ngân hàng (Vietcombank - STK: 0011004567890)',
  },
  {
    key: 'Kem_Theo',
    label: 'Chứng từ kèm theo',
    category: 'Tài chính & Văn bản',
    placeholder: '02 Hóa đơn GTGT điện tử và Biên bản nghiệm thu dịch vụ.',
  },
  {
    key: 'Thiet_Bi_Tai_San',
    label: 'Tên thiết bị / Tài sản',
    category: 'Tài chính & Văn bản',
    placeholder: 'Máy tính xách tay MacBook Pro 14 inch M3 (S/N: C02X9876ABCD)',
  },
  {
    key: 'Tinh_Trang',
    label: 'Tình trạng tài sản / Thiết bị',
    category: 'Tài chính & Văn bản',
    placeholder: 'Mới 100%, hoạt động tốt, kèm bộ sạc 70W nguyên hộp.',
  },
  {
    key: 'Ghi_Chu_Dau_Trang',
    label: 'Tiêu đề đầu trang (Headnote / Header)',
    category: 'Tài chính & Văn bản',
    placeholder: 'BIỂU MẪU HÀNH CHÍNH CHUẨN ISO 9001:2015 — MÃ SỐ: BM-TCKT-01',
  },
  {
    key: 'Chu_Thich_Footnote',
    label: 'Chú thích chân trang (Footnote)',
    category: 'Tài chính & Văn bản',
    placeholder: 'Chú thích (1): Đề nghị đính kèm đầy đủ hóa đơn GTGT và biên bản nghiệm thu khi trình ký.',
  },
];

export interface FieldSymbolPair {
  openSymbol: string;
  closeSymbol: string;
}

export interface FieldSymbolPreset extends FieldSymbolPair {
  id: string;
  name: string;
  example: string;
}

export const PRESET_FIELD_SYMBOLS: FieldSymbolPreset[] = [
  {
    id: 'guillemet',
    name: 'Chuẩn Mailings Word «...»',
    openSymbol: '«',
    closeSymbol: '»',
    example: '«Ho_Va_Ten»',
  },
  {
    id: 'double_curly',
    name: 'Ngoặc nhọn kép {{...}}',
    openSymbol: '{{',
    closeSymbol: '}}',
    example: '{{Ho_Va_Ten}}',
  },
  {
    id: 'single_curly',
    name: 'Ngoặc nhọn đơn {...}',
    openSymbol: '{',
    closeSymbol: '}',
    example: '{Ho_Va_Ten}',
  },
  {
    id: 'double_square',
    name: 'Ngoặc vuông kép [[...]]',
    openSymbol: '[[',
    closeSymbol: ']]',
    example: '[[Ho_Va_Ten]]',
  },
  {
    id: 'single_square',
    name: 'Ngoặc vuông [...]',
    openSymbol: '[',
    closeSymbol: ']',
    example: '[Ho_Va_Ten]',
  },
  {
    id: 'double_angle',
    name: 'Mũi tên kép <<...>>',
    openSymbol: '<<',
    closeSymbol: '>>',
    example: '<<Ho_Va_Ten>>',
  },
  {
    id: 'dollar_curly',
    name: 'Biến ${...}',
    openSymbol: '${',
    closeSymbol: '}',
    example: '${Ho_Va_Ten}',
  },
];

export const DEFAULT_FIELD_SYMBOL: FieldSymbolPair = {
  openSymbol: '«',
  closeSymbol: '»',
};

export interface SavedTemplateEntry {
  file_path: string;
  selected_indices: string[];
  custom_labels?: Record<string, string>;
  paragraph_templates?: Record<string, string>;
  mailings_fields?: Record<
    string,
    {
      label: string;
      defaultValue: string;
      category?: MailingsFieldCategory;
    }
  >;
  field_symbol?: FieldSymbolPair;
  base64_docx: string;
  updated_at: string;
  total_paragraphs: number;
  category?: string;
  drive_file_id?: string;
  drive_web_view_link?: string;
}

export type TemplatesConfigMap = Record<string, SavedTemplateEntry>;

export interface SharedDrivePublicationInfo {
  backupFileId: string;
  backupFileName: string;
  backupWebViewLink?: string;
  folderId?: string;
  folderName?: string;
  folderWebViewLink?: string;
  ownerEmail?: string;
  updatedAt: string;
  templateCount: number;
}

export interface ExportHistoryRecord {
  id: string;
  templateName: string;
  outputFileName: string;
  exportedAt: string;
  modifiedCount: number;
  replacements: Record<number, string>;
  fieldValues?: Record<string, string>;
}

const STORAGE_KEY_CONFIG = 'smart_word_templates_config_v2';
const STORAGE_KEY_HISTORY = 'smart_word_export_history_v2';
const STORAGE_KEY_SEEDED = 'smart_word_templates_seeded_v4';
const STORAGE_KEY_CUSTOM_FIELDS = 'smart_word_custom_mailings_fields_v2';
const STORAGE_KEY_PALETTE_OVERRIDES = 'smart_word_palette_fields_v1';
const STORAGE_KEY_FIELD_SYMBOL = 'smart_word_field_symbol_v1';

export function loadFieldSymbolConfig(): FieldSymbolPair {
  try {
    const raw = localStorage.getItem(STORAGE_KEY_FIELD_SYMBOL);
    if (!raw) return DEFAULT_FIELD_SYMBOL;
    const parsed = JSON.parse(raw) as FieldSymbolPair;
    if (parsed && parsed.openSymbol && parsed.closeSymbol) {
      return parsed;
    }
    return DEFAULT_FIELD_SYMBOL;
  } catch {
    return DEFAULT_FIELD_SYMBOL;
  }
}

export function saveFieldSymbolConfig(symbol: FieldSymbolPair): void {
  try {
    localStorage.setItem(STORAGE_KEY_FIELD_SYMBOL, JSON.stringify(symbol));
  } catch {
    // Ignore storage error
  }
}

export function formatFieldToken(
  key: string,
  symbol: FieldSymbolPair = DEFAULT_FIELD_SYMBOL
): string {
  return `${symbol.openSymbol}${key}${symbol.closeSymbol}`;
}

export function normalizeFieldKeyInput(raw: string): string {
  const cleaned = raw
    .trim()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/đ/g, 'd')
    .replace(/Đ/g, 'D')
    .replace(/[^a-zA-Z0-9_]+/g, '_')
    .replace(/^_+|_+$/g, '');
  if (/^\d+$/.test(cleaned)) {
    return `Truong_${cleaned}`;
  }
  return cleaned;
}

function escapeRegExp(str: string): string {
  return str.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

/**
 * Valid Mailings field key pattern: must contain at least one letter so purely numeric
 * strings like "04" or "[04]" (e.g. footnote numbers, list indices) are never falsely treated as merge fields.
 */
const VALID_FIELD_KEY_PATTERN = '[a-zA-Z_]*[a-zA-Z][a-zA-Z0-9_]*';

/**
 * Builds a regex that matches standard multi-character field delimiters («Key», {{Key}}, [[Key]], <<Key>>, ${Key})
 * plus any custom openSymbol/closeSymbol configured by the user.
 * Single brackets [Key] or {Key} are only matched when explicitly selected as customSymbol.
 */
export function buildUniversalFieldTokenRegex(customSymbol?: FieldSymbolPair, globalFlag = true): RegExp {
  const patterns: string[] = [
    `«(${VALID_FIELD_KEY_PATTERN})»`,
    `\\{\\{(${VALID_FIELD_KEY_PATTERN})\\}\\}`,
    `\\[\\[(${VALID_FIELD_KEY_PATTERN})\\]\\]`,
    `<<(${VALID_FIELD_KEY_PATTERN})>>`,
    `\\$\\{(${VALID_FIELD_KEY_PATTERN})\\}`,
  ];

  if (
    customSymbol &&
    customSymbol.openSymbol &&
    customSymbol.closeSymbol
  ) {
    const customPattern = `${escapeRegExp(customSymbol.openSymbol)}(${VALID_FIELD_KEY_PATTERN})${escapeRegExp(
      customSymbol.closeSymbol
    )}`;
    if (!patterns.includes(customPattern)) {
      patterns.unshift(customPattern);
    }
  }

  return new RegExp(patterns.join('|'), globalFlag ? 'g' : '');
}

/**
 * Builds a regex targeting a specific field key across all possible delimiters
 * (including numeric keys like "04" or single brackets "[04]" so legacy/stuck tokens can always be removed or swapped).
 */
function buildSpecificKeyTokenRegex(targetKey: string, symbol?: FieldSymbolPair): RegExp {
  const escapedKey = escapeRegExp(targetKey);
  const tokenPatterns: string[] = [
    `«${escapedKey}»`,
    `\\{\\{${escapedKey}\\}\\}`,
    `\\[\\[${escapedKey}\\]\\]`,
    `<<${escapedKey}>>`,
    `\\$\\{${escapedKey}\\}`,
    `\\{${escapedKey}\\}`,
    `\\[${escapedKey}\\]`,
  ];

  if (symbol && symbol.openSymbol && symbol.closeSymbol) {
    const customToken = `${escapeRegExp(symbol.openSymbol)}${escapedKey}${escapeRegExp(symbol.closeSymbol)}`;
    if (!tokenPatterns.includes(customToken)) {
      tokenPatterns.unshift(customToken);
    }
  }

  return new RegExp(tokenPatterns.join('|'), 'g');
}

/**
 * Extracts the field key if `tokenPart` is a complete field token (e.g. "«Ho_Va_Ten»" -> "Ho_Va_Ten").
 */
export function matchSingleFieldToken(
  tokenPart: string,
  customSymbol?: FieldSymbolPair
): string | null {
  const singleRegex = new RegExp(
    `^(?:${buildUniversalFieldTokenRegex(customSymbol, false).source})$`
  );
  const m = tokenPart.match(singleRegex);
  if (!m) return null;
  for (let i = 1; i < m.length; i++) {
    if (m[i]) return m[i];
  }
  return null;
}

export function loadCustomMailingsFields(): MailingsFieldDef[] {
  try {
    const rawPalette = localStorage.getItem(STORAGE_KEY_PALETTE_OVERRIDES);
    if (rawPalette) {
      const parsed = JSON.parse(rawPalette) as MailingsFieldDef[];
      if (Array.isArray(parsed) && parsed.length > 0) {
        return parsed;
      }
    }
    const rawCustom = localStorage.getItem(STORAGE_KEY_CUSTOM_FIELDS);
    const legacyCustom = rawCustom ? (JSON.parse(rawCustom) as MailingsFieldDef[]) : [];
    const merged = [...DEFAULT_MAILINGS_FIELDS];
    legacyCustom.forEach((cf) => {
      const idx = merged.findIndex((d) => d.key === cf.key);
      if (idx >= 0) {
        merged[idx] = cf;
      } else {
        merged.push(cf);
      }
    });
    return merged;
  } catch {
    return [...DEFAULT_MAILINGS_FIELDS];
  }
}

export function saveCustomMailingsFields(fields: MailingsFieldDef[]): void {
  localStorage.setItem(STORAGE_KEY_PALETTE_OVERRIDES, JSON.stringify(fields));
  localStorage.setItem(STORAGE_KEY_CUSTOM_FIELDS, JSON.stringify(fields));
}

export function resetDefaultMailingsFields(): MailingsFieldDef[] {
  localStorage.removeItem(STORAGE_KEY_PALETTE_OVERRIDES);
  localStorage.removeItem(STORAGE_KEY_CUSTOM_FIELDS);
  return [...DEFAULT_MAILINGS_FIELDS];
}

export function extractMergeFieldKeysFromText(
  text: string,
  customSymbol?: FieldSymbolPair
): string[] {
  if (!text) return [];
  const keys: string[] = [];
  const regex = buildUniversalFieldTokenRegex(customSymbol, true);
  let match: RegExpExecArray | null;
  while ((match = regex.exec(text)) !== null) {
    let foundKey: string | undefined;
    for (let i = 1; i < match.length; i++) {
      if (match[i]) {
        foundKey = match[i];
        break;
      }
    }
    if (foundKey && !keys.includes(foundKey)) {
      keys.push(foundKey);
    }
  }
  return keys;
}

export function hasAnyMergeFieldToken(
  text: string,
  customSymbol?: FieldSymbolPair
): boolean {
  return extractMergeFieldKeysFromText(text, customSymbol).length > 0;
}

export function resolveParagraphTemplate(
  templateStr: string,
  fieldValues: Record<string, string>,
  keepTokenIfEmpty = false,
  customSymbol?: FieldSymbolPair
): string {
  const regex = buildUniversalFieldTokenRegex(customSymbol, true);
  return templateStr.replace(regex, (fullMatch, ...args) => {
    let key: string | undefined;
    for (let i = 0; i < args.length - 2; i++) {
      if (typeof args[i] === 'string' && args[i]) {
        key = args[i];
        break;
      }
    }
    if (!key) return fullMatch;
    const val = fieldValues[key];
    if (val === undefined || (keepTokenIfEmpty && val === '')) {
      return fullMatch;
    }
    return val;
  });
}

/**
 * Replaces all occurrences of `oldKey` token in `templateStr` with `newKey` formatted using `symbol`.
 */
export function replaceFieldKeyInTemplateString(
  templateStr: string,
  oldKey: string,
  newKey: string,
  symbol: FieldSymbolPair = DEFAULT_FIELD_SYMBOL
): string {
  if (!templateStr || !oldKey) return templateStr;
  const specificRegex = buildSpecificKeyTokenRegex(oldKey, symbol);
  return templateStr.replace(specificRegex, formatFieldToken(newKey, symbol));
}

/**
 * Removes all occurrences of `targetKey` token in `templateStr` across any delimiter («Key», {{Key}}, [Key], etc.).
 */
export function removeFieldKeyFromTemplateString(
  templateStr: string,
  targetKey: string,
  symbol: FieldSymbolPair = DEFAULT_FIELD_SYMBOL
): string {
  if (!templateStr || !targetKey) return templateStr;
  const specificRegex = buildSpecificKeyTokenRegex(targetKey, symbol);
  return templateStr
    .replace(specificRegex, '')
    .replace(/\s{2,}/g, ' ')
    .trim();
}

/**
 * Converts all field tokens in `templateStr` to use the specified `targetSymbol` (e.g. «Key» -> {{Key}}).
 */
export function convertTemplateTokensToSymbol(
  templateStr: string,
  targetSymbol: FieldSymbolPair,
  previousCustomSymbol?: FieldSymbolPair
): string {
  if (!templateStr) return templateStr;
  const regex = buildUniversalFieldTokenRegex(previousCustomSymbol || targetSymbol, true);
  return templateStr.replace(regex, (fullMatch, ...args) => {
    let matchedKey: string | undefined;
    for (let i = 0; i < args.length - 2; i++) {
      if (typeof args[i] === 'string' && args[i]) {
        matchedKey = args[i];
        break;
      }
    }
    if (!matchedKey) return fullMatch;
    return formatFieldToken(matchedKey, targetSymbol);
  });
}

export function getFieldDefByKey(
  key: string,
  customFields: MailingsFieldDef[] = []
): MailingsFieldDef {
  const found =
    customFields.find((f) => f.key === key) ||
    DEFAULT_MAILINGS_FIELDS.find((f) => f.key === key);
  if (found) return found;
  return {
    key,
    label: key.replace(/_/g, ' '),
    category: 'Trường tùy chỉnh',
    placeholder: `Nhập ${key.replace(/_/g, ' ').toLowerCase()}...`,
  };
}

export function loadTemplatesConfig(): TemplatesConfigMap {
  try {
    const raw = localStorage.getItem(STORAGE_KEY_CONFIG);
    if (!raw) return {};
    return JSON.parse(raw) as TemplatesConfigMap;
  } catch {
    return {};
  }
}

/**
 * Saves templates config to both localStorage AND the shared backend server (`templates_store/config.json`)
 * so that anyone visiting the shared app URL immediately has access to all saved templates.
 */
export function saveTemplatesConfig(config: TemplatesConfigMap): void {
  try {
    localStorage.setItem(STORAGE_KEY_CONFIG, JSON.stringify(config, null, 2));
  } catch {
    // Ignore localStorage quota errors
  }

  fetch('/api/templates', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ templates: config, mode: 'replace' }),
  }).catch((err) => {
    console.warn('Could not sync templates to server:', err);
  });
}

/**
 * Fetches shared templates from the server (`/api/templates`) and merges with local templates.
 */
export async function syncTemplatesWithServer(): Promise<TemplatesConfigMap> {
  const localConfig = loadTemplatesConfig();
  try {
    const res = await fetch('/api/templates');
    if (res.ok) {
      const data = await res.json();
      const serverConfig = (data?.templates || {}) as TemplatesConfigMap;
      const serverKeys = Object.keys(serverConfig);
      const localKeys = Object.keys(localConfig);

      if (serverKeys.length > 0 || localKeys.length > 0) {
        // Merge so templates created locally or on server are both preserved
        const merged: TemplatesConfigMap = { ...localConfig, ...serverConfig };
        // If local had newer updates or extra templates not on server yet, push merged back to server
        let needsPush = false;
        for (const k of localKeys) {
          if (!serverConfig[k]) {
            merged[k] = localConfig[k];
            needsPush = true;
          } else if (
            (localConfig[k].updated_at || '') > (serverConfig[k].updated_at || '')
          ) {
            merged[k] = localConfig[k];
            needsPush = true;
          }
        }

        try {
          localStorage.setItem(STORAGE_KEY_CONFIG, JSON.stringify(merged, null, 2));
        } catch {
          // Ignore
        }

        if (needsPush) {
          await fetch('/api/templates', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ templates: merged, mode: 'merge' }),
          });
        }

        return merged;
      }
    }
  } catch (err) {
    console.warn('Failed to fetch shared templates from server:', err);
  }
  return localConfig;
}

/**
 * Encodes a single template entry into a shareable URL string (`#share_tpl=...`)
 * so that even if a link is sent to someone, opening the URL imports that template directly.
 */
export function buildShareableTemplateUrl(
  templateName: string,
  entry: SavedTemplateEntry
): string {
  const baseUrl = window.location.origin + window.location.pathname;
  try {
    const payload = JSON.stringify({ name: templateName, entry });
    const encoded = window.btoa(unescape(encodeURIComponent(payload)));
    return `${baseUrl}#share_tpl=${encodeURIComponent(encoded)}`;
  } catch {
    return baseUrl;
  }
}

/**
 * Checks if the current URL hash contains `#share_tpl=...` and returns the decoded template if present.
 */
export function extractSharedTemplateFromUrlHash(): {
  name: string;
  entry: SavedTemplateEntry;
} | null {
  try {
    const hash = window.location.hash;
    if (!hash || !hash.startsWith('#share_tpl=')) return null;
    const encoded = decodeURIComponent(hash.slice('#share_tpl='.length));
    const jsonStr = decodeURIComponent(escape(window.atob(encoded)));
    const parsed = JSON.parse(jsonStr);
    if (parsed && typeof parsed.name === 'string' && parsed.entry?.base64_docx) {
      // Clean the URL hash so refreshing doesn't re-trigger
      window.history.replaceState(null, '', window.location.pathname + window.location.search);
      return { name: parsed.name, entry: parsed.entry as SavedTemplateEntry };
    }
  } catch (err) {
    console.warn('Could not decode shared template from URL hash:', err);
  }
  return null;
}

/**
 * Checks if the URL contains `?driveId=<FILE_ID>` so that opening a link pointing to a shared Google Drive
 * backup file automatically loads that Google Drive library for any visitor.
 */
export function extractDriveIdFromUrlQuery(): string | null {
  try {
    const params = new URLSearchParams(window.location.search);
    const driveId = params.get('driveId');
    if (driveId && /^[a-zA-Z0-9_-]+$/.test(driveId.trim())) {
      return driveId.trim();
    }
  } catch {
    // Ignore
  }
  return null;
}

/**
 * Fetches the currently configured shared Google Drive info + templates from the server.
 * When `pullLatestFromDrive = true`, the server directly downloads the latest shared JSON file from Google Drive.
 */
export async function fetchSharedDriveConfig(
  pullLatestFromDrive = false
): Promise<{
  driveConfig: SharedDrivePublicationInfo | null;
  templates: TemplatesConfigMap;
  syncedFromDrive: boolean;
}> {
  try {
    const res = await fetch(
      `/api/drive-config${pullLatestFromDrive ? '?pull=true' : ''}`
    );
    if (res.ok) {
      const data = await res.json();
      const templates = (data?.templates || {}) as TemplatesConfigMap;
      if (Object.keys(templates).length > 0) {
        try {
          localStorage.setItem(STORAGE_KEY_CONFIG, JSON.stringify(templates, null, 2));
        } catch {
          // Ignore
        }
      }
      return {
        driveConfig: (data?.driveConfig || null) as SharedDrivePublicationInfo | null,
        templates,
        syncedFromDrive: Boolean(data?.syncedFromDrive),
      };
    }
  } catch (err) {
    console.warn('Could not fetch shared Google Drive config:', err);
  }
  return {
    driveConfig: null,
    templates: loadTemplatesConfig(),
    syncedFromDrive: false,
  };
}

/**
 * Links any publicly shared Google Drive file ID (`templates_store_backup.json`) so all app visitors
 * load templates directly from that Google Drive file.
 */
export async function loadPublicDriveLibraryByFileId(fileId: string): Promise<{
  driveConfig: SharedDrivePublicationInfo;
  templates: TemplatesConfigMap;
}> {
  const res = await fetch(`/api/drive-public/${encodeURIComponent(fileId)}`);
  const data = await res.json().catch(() => ({}));
  if (!res.ok) {
    throw new Error(
      data?.error ||
        'Không thể tải dữ liệu từ Google Drive. Hãy kiểm tra quyền chia sẻ công khai của tệp trên Google Drive.'
    );
  }
  const templates = (data.templates || {}) as TemplatesConfigMap;
  try {
    localStorage.setItem(STORAGE_KEY_CONFIG, JSON.stringify(templates, null, 2));
  } catch {
    // Ignore
  }
  return {
    driveConfig: data.driveConfig as SharedDrivePublicationInfo,
    templates,
  };
}

export function loadExportHistory(): ExportHistoryRecord[] {
  try {
    const raw = localStorage.getItem(STORAGE_KEY_HISTORY);
    if (!raw) return [];
    return JSON.parse(raw) as ExportHistoryRecord[];
  } catch {
    return [];
  }
}

export async function syncHistoryWithServer(): Promise<ExportHistoryRecord[]> {
  const local = loadExportHistory();
  try {
    const res = await fetch('/api/history');
    if (res.ok) {
      const data = await res.json();
      const serverHist = (data?.history || []) as ExportHistoryRecord[];
      if (serverHist.length > 0) {
        localStorage.setItem(STORAGE_KEY_HISTORY, JSON.stringify(serverHist));
        return serverHist;
      }
    }
  } catch {
    // Ignore
  }
  return local;
}

export function addExportHistory(record: ExportHistoryRecord): ExportHistoryRecord[] {
  const current = loadExportHistory();
  const next = [record, ...current].slice(0, 30);
  try {
    localStorage.setItem(STORAGE_KEY_HISTORY, JSON.stringify(next));
  } catch {
    // Ignore
  }
  fetch('/api/history', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ record }),
  }).catch(() => {});
  return next;
}

export function clearExportHistory(): void {
  localStorage.removeItem(STORAGE_KEY_HISTORY);
  fetch('/api/history', { method: 'DELETE' }).catch(() => {});
}

export function formatPythonCompatibleConfig(config: TemplatesConfigMap): string {
  const pythonClean: Record<
    string,
    {
      file_path: string;
      selected_indices: string[];
      paragraph_templates?: Record<string, string>;
    }
  > = {};
  Object.entries(config).forEach(([name, entry]) => {
    pythonClean[name] = {
      file_path: entry.file_path,
      selected_indices: entry.selected_indices,
      ...(entry.paragraph_templates ? { paragraph_templates: entry.paragraph_templates } : {}),
    };
  });
  return JSON.stringify(pythonClean, null, 4);
}

const SAMPLE_PAYMENT_SPEC: SampleParagraphSpec[] = [
  { text: 'CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM', align: 'center', bold: true, fontSizeHalfPt: 26 },
  { text: 'Độc lập - Tự do - Hạnh phúc', align: 'center', bold: true, fontSizeHalfPt: 26 },
  { text: '--------------------------------', align: 'center', fontSizeHalfPt: 24 },
  { text: 'Hà Nội, ngày 23 tháng 10 năm 2026', align: 'right', italic: true, fontSizeHalfPt: 26 },
  { text: 'GIẤY ĐỀ NGHỊ THANH TOÁN', align: 'center', bold: true, fontSizeHalfPt: 30 },
  { text: 'Kính gửi: Ban Giám đốc và Phòng Tài chính - Kế toán', align: 'center', bold: true, fontSizeHalfPt: 26 },
  { text: 'Họ và tên người đề nghị: Nguyễn Văn Minh', align: 'left', fontSizeHalfPt: 26 },
  { text: 'Bộ phận / Phòng ban: Phòng Triển khai Dự án Công nghệ', align: 'left', fontSizeHalfPt: 26 },
  { text: 'Nội dung thanh toán: Chi phí mua bản quyền máy chủ và tên miền quý IV/2026', align: 'left', fontSizeHalfPt: 26 },
  { text: 'Số tiền đề nghị thanh toán: 18.500.000 VNĐ', align: 'left', bold: true, fontSizeHalfPt: 26, footnoteRefId: 1 },
  { text: 'Số tiền bằng chữ: Mười tám triệu năm trăm nghìn đồng chẵn.', align: 'left', italic: true, fontSizeHalfPt: 26 },
  { text: 'Hình thức thanh toán: Chuyển khoản ngân hàng (Vietcombank - STK: 0011004567890)', align: 'left', fontSizeHalfPt: 26 },
  { text: 'Kèm theo: 02 Hóa đơn GTGT điện tử và Biên bản nghiệm thu dịch vụ.', align: 'left', fontSizeHalfPt: 26 },
  { text: 'Người đề nghị ký tên: Nguyễn Văn Minh', align: 'right', bold: true, fontSizeHalfPt: 26 },
];

const SAMPLE_LEAVE_SPEC: SampleParagraphSpec[] = [
  { text: 'CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM', align: 'center', bold: true, fontSizeHalfPt: 26 },
  { text: 'Độc lập - Tự do - Hạnh phúc', align: 'center', bold: true, fontSizeHalfPt: 26 },
  { text: 'ĐƠN XIN NGHỈ PHÉP NĂM', align: 'center', bold: true, fontSizeHalfPt: 30 },
  { text: 'Kính gửi: Trưởng phòng Nhân sự và Quản lý trực tiếp', align: 'center', fontSizeHalfPt: 26 },
  { text: 'Tôi tên là: Trần Thị Thu Hà', align: 'left', fontSizeHalfPt: 26 },
  { text: 'Mã nhân viên: NV-2048 | Chức vụ: Chuyên viên Phân tích Nghiệp vụ', align: 'left', fontSizeHalfPt: 26 },
  { text: 'Thời gian xin nghỉ: Từ ngày 26/10/2026 đến hết ngày 28/10/2026 (03 ngày)', align: 'left', bold: true, fontSizeHalfPt: 26, footnoteRefId: 1 },
  { text: 'Lý do xin nghỉ: Giải quyết công việc gia đình tại quê nhà.', align: 'left', fontSizeHalfPt: 26 },
  { text: 'Người nhận bàn giao công việc trong thời gian nghỉ: Lê Hoàng Nam (Phòng BA)', align: 'left', fontSizeHalfPt: 26 },
  { text: 'Số điện thoại liên hệ khẩn cấp: 0912.345.678', align: 'left', fontSizeHalfPt: 26 },
  { text: 'Tôi xin cam đoan hoàn thành đầy đủ tiến độ công việc trước khi nghỉ phép.', align: 'left', fontSizeHalfPt: 26 },
  { text: 'TP. Hồ Chí Minh, ngày 23 tháng 10 năm 2026', align: 'right', italic: true, fontSizeHalfPt: 26 },
];

const SAMPLE_HANDOVER_SPEC: SampleParagraphSpec[] = [
  { text: 'CÔNG TY CỔ PHẦN GIẢI PHÁP SỐ ĐÔNG DƯƠNG', align: 'center', bold: true, fontSizeHalfPt: 26 },
  { text: 'Số: 42/2026/BBBG-HCNS', align: 'center', fontSizeHalfPt: 24 },
  { text: 'BIÊN BẢN BÀN GIAO TÀI SẢN & THIẾT BỊ LÀM VIỆC', align: 'center', bold: true, fontSizeHalfPt: 30 },
  { text: 'Hôm nay, ngày 23 tháng 10 năm 2026, tại Văn phòng Công ty, chúng tôi gồm:', align: 'left', fontSizeHalfPt: 26 },
  { text: 'Bên giao (Phòng IT - Hành chính): Ông Phạm Quốc Khánh - Chuyên viên Quản trị Hệ thống', align: 'left', fontSizeHalfPt: 26 },
  { text: 'Bên nhận (Nhân sự tiếp nhận): Bà Vũ Ngọc Mai - Phòng Kinh doanh Doanh nghiệp', align: 'left', fontSizeHalfPt: 26 },
  { text: 'Thiết bị bàn giao: Máy tính xách tay MacBook Pro 14 inch M3 (S/N: C02X9876ABCD)', align: 'left', bold: true, fontSizeHalfPt: 26, footnoteRefId: 1 },
  { text: 'Tình trạng thiết bị khi bàn giao: Mới 100%, hoạt động tốt, kèm bộ sạc 70W nguyên hộp.', align: 'left', fontSizeHalfPt: 26 },
  { text: 'Bên nhận có trách nhiệm bảo quản, sử dụng thiết bị đúng mục đích công việc của Công ty.', align: 'left', fontSizeHalfPt: 26 },
  { text: 'Biên bản được lập thành 02 bản có giá trị pháp lý như nhau, mỗi bên giữ 01 bản.', align: 'left', italic: true, fontSizeHalfPt: 26 },
];

export async function ensureStarterTemplatesSeeded(): Promise<TemplatesConfigMap> {
  // First check shared server templates + local templates
  const synced = await syncTemplatesWithServer();
  const hasV3Seeded = localStorage.getItem(STORAGE_KEY_SEEDED) === 'true';

  // Check if default starter templates still have old format without Header (Headnote) & Footnote
  const needsStarterUpgrade =
    !hasV3Seeded ||
    (synced['Giấy đề nghị thanh toán'] &&
      synced['Giấy đề nghị thanh toán'].total_paragraphs <= 14);

  if (Object.keys(synced).length > 0 && !needsStarterUpgrade) {
    return synced;
  }

  // Note on paragraph indices when generated by generateSampleDocxBase64:
  // Index 0: Header (Headnote / Đầu trang)
  // Index 1..N: Body paragraphs (shifted by +1 because Header is at index 0)
  // Index N+1: Footnote #1 (Chú thích chân trang)
  // Index N+2: Footer (Chân trang)
  const [paymentBase64, leaveBase64, handoverBase64] = await Promise.all([
    generateSampleDocxBase64(SAMPLE_PAYMENT_SPEC, {
      headerLines: [
        {
          text: 'BIỂU MẪU HÀNH CHÍNH CHUẨN ISO 9001:2015 — MÃ SỐ: BM-TCKT-01',
          align: 'right',
          italic: true,
          fontSizeHalfPt: 20,
        },
      ],
      footnotes: [
        {
          id: 1,
          text: 'Chú thích (1): Đề nghị đính kèm đầy đủ hóa đơn GTGT và biên bản nghiệm thu khi trình ký.',
        },
      ],
      footerLines: [
        {
          text: 'Tài liệu lưu hành nội bộ — Phòng Tài chính Kế toán · Trang 1/1',
          align: 'center',
          italic: true,
          fontSizeHalfPt: 20,
        },
      ],
    }),
    generateSampleDocxBase64(SAMPLE_LEAVE_SPEC, {
      headerLines: [
        {
          text: 'BIỂU MẪU NHÂN SỰ — MÃ SỐ: BM-HCNS-02',
          align: 'right',
          italic: true,
          fontSizeHalfPt: 20,
        },
      ],
      footnotes: [
        {
          id: 1,
          text: 'Chú thích (1): Đơn xin nghỉ phép từ 03 ngày trở lên cần có xác nhận bàn giao công việc của Trưởng bộ phận.',
        },
      ],
      footerLines: [
        {
          text: 'Hồ sơ Quản trị Nhân sự · Trang 1/1',
          align: 'center',
          italic: true,
          fontSizeHalfPt: 20,
        },
      ],
    }),
    generateSampleDocxBase64(SAMPLE_HANDOVER_SPEC, {
      headerLines: [
        {
          text: 'BIỂU MẪU QUẢN TRỊ THIẾT BỊ & TÀI SẢN — MÃ SỐ: BM-QTTS-03',
          align: 'right',
          italic: true,
          fontSizeHalfPt: 20,
        },
      ],
      footnotes: [
        {
          id: 1,
          text: 'Chú thích (1): Kiểm tra kỹ số Serial Number (S/N) và tem bảo hành trước khi ký nhận bàn giao.',
        },
      ],
      footerLines: [
        {
          text: 'Biên bản lưu tại Phòng IT & Hành chính · Trang 1/1',
          align: 'center',
          italic: true,
          fontSizeHalfPt: 20,
        },
      ],
    }),
  ]);

  const starterConfig: TemplatesConfigMap = {
    'Giấy đề nghị thanh toán': {
      file_path: 'templates_store/giay_de_nghi_thanh_toan.docx',
      selected_indices: ['0', '4', '6', '7', '8', '9', '10', '11', '12', '13', '14', '15'],
      paragraph_templates: {
        '0': '«Ghi_Chu_Dau_Trang»',
        '4': '«Dia_Diem», ngày «Ngay» tháng «Thang» năm «Nam»',
        '6': 'Kính gửi: «Kinh_Gui»',
        '7': 'Họ và tên người đề nghị: «Ho_Va_Ten»',
        '8': 'Bộ phận / Phòng ban: «Phong_Ban»',
        '9': 'Nội dung thanh toán: «Noi_Dung»',
        '10': 'Số tiền đề nghị thanh toán: «So_Tien_So»',
        '11': 'Số tiền bằng chữ: «So_Tien_Chu»',
        '12': 'Hình thức thanh toán: «Hinh_Thuc_Thanh_Toan»',
        '13': 'Kèm theo: «Kem_Theo»',
        '14': 'Người đề nghị ký tên: «Ho_Va_Ten»',
        '15': '«Chu_Thich_Footnote»',
      },
      mailings_fields: {
        Ghi_Chu_Dau_Trang: { label: 'Tiêu đề đầu trang (Headnote / Header)', defaultValue: 'BIỂU MẪU HÀNH CHÍNH CHUẨN ISO 9001:2015 — MÃ SỐ: BM-TCKT-01', category: 'Tài chính & Văn bản' },
        Dia_Diem: { label: 'Địa điểm lập văn bản', defaultValue: 'Hà Nội', category: 'Thời gian & Địa điểm' },
        Ngay: { label: 'Ngày', defaultValue: '23', category: 'Thời gian & Địa điểm' },
        Thang: { label: 'Tháng', defaultValue: '10', category: 'Thời gian & Địa điểm' },
        Nam: { label: 'Năm', defaultValue: '2026', category: 'Thời gian & Địa điểm' },
        Kinh_Gui: { label: 'Kính gửi', defaultValue: 'Ban Giám đốc và Phòng Tài chính - Kế toán', category: 'Đơn vị & Tổ chức' },
        Ho_Va_Ten: { label: 'Họ và tên', defaultValue: 'Nguyễn Văn Minh', category: 'Cá nhân & Nhân sự' },
        Phong_Ban: { label: 'Phòng ban / Bộ phận', defaultValue: 'Phòng Triển khai Dự án Công nghệ', category: 'Cá nhân & Nhân sự' },
        Noi_Dung: { label: 'Nội dung thanh toán', defaultValue: 'Chi phí mua bản quyền máy chủ và tên miền quý IV/2026', category: 'Tài chính & Văn bản' },
        So_Tien_So: { label: 'Số tiền bằng số', defaultValue: '18.500.000 VNĐ', category: 'Tài chính & Văn bản' },
        So_Tien_Chu: { label: 'Số tiền bằng chữ', defaultValue: 'Mười tám triệu năm trăm nghìn đồng chẵn.', category: 'Tài chính & Văn bản' },
        Hinh_Thuc_Thanh_Toan: { label: 'Hình thức & Số tài khoản', defaultValue: 'Chuyển khoản ngân hàng (Vietcombank - STK: 0011004567890)', category: 'Tài chính & Văn bản' },
        Kem_Theo: { label: 'Chứng từ kèm theo', defaultValue: '02 Hóa đơn GTGT điện tử và Biên bản nghiệm thu dịch vụ.', category: 'Tài chính & Văn bản' },
        Chu_Thich_Footnote: { label: 'Chú thích chân trang (Footnote)', defaultValue: 'Chú thích (1): Đề nghị đính kèm đầy đủ hóa đơn GTGT và biên bản nghiệm thu khi trình ký.', category: 'Tài chính & Văn bản' },
      },
      custom_labels: {
        '0': 'Đầu trang (Headnote / Header)',
        '4': 'Địa điểm & Ngày tháng',
        '6': 'Kính gửi',
        '7': 'Họ và tên người đề nghị',
        '8': 'Bộ phận / Phòng ban',
        '9': 'Nội dung thanh toán',
        '10': 'Số tiền bằng số (Có gắn Footnote #1)',
        '11': 'Số tiền bằng chữ',
        '12': 'Thông tin tài khoản nhận',
        '13': 'Chứng từ kèm theo',
        '14': 'Chữ ký người đề nghị (Tự động đồng bộ Họ tên)',
        '15': 'Chú thích chân trang (Footnote #1)',
      },
      base64_docx: paymentBase64,
      updated_at: '2026-10-23T06:00:00.000Z',
      total_paragraphs: SAMPLE_PAYMENT_SPEC.length + 3,
      category: 'Tài chính - Kế toán',
    },
    'Đơn xin nghỉ phép năm': {
      file_path: 'templates_store/don_xin_nghi_phep_nam.docx',
      selected_indices: ['0', '4', '5', '6', '7', '8', '9', '10', '12', '13'],
      paragraph_templates: {
        '0': '«Ghi_Chu_Dau_Trang»',
        '4': 'Kính gửi: «Kinh_Gui»',
        '5': 'Tôi tên là: «Ho_Va_Ten»',
        '6': 'Mã nhân viên: «Ma_Nhan_Vien» | Chức vụ: «Chuc_Vu»',
        '7': 'Thời gian xin nghỉ: «Thoi_Gian_Tu_Den»',
        '8': 'Lý do xin nghỉ: «Ly_Do»',
        '9': 'Người nhận bàn giao công việc trong thời gian nghỉ: «Ben_Nhan»',
        '10': 'Số điện thoại liên hệ khẩn cấp: «So_Dien_Thoai»',
        '12': '«Dia_Diem», ngày «Ngay» tháng «Thang» năm «Nam»',
        '13': '«Chu_Thich_Footnote»',
      },
      mailings_fields: {
        Ghi_Chu_Dau_Trang: { label: 'Tiêu đề đầu trang (Headnote / Header)', defaultValue: 'BIỂU MẪU NHÂN SỰ — MÃ SỐ: BM-HCNS-02', category: 'Tài chính & Văn bản' },
        Kinh_Gui: { label: 'Kính gửi', defaultValue: 'Trưởng phòng Nhân sự và Quản lý trực tiếp', category: 'Đơn vị & Tổ chức' },
        Ho_Va_Ten: { label: 'Họ và tên', defaultValue: 'Trần Thị Thu Hà', category: 'Cá nhân & Nhân sự' },
        Ma_Nhan_Vien: { label: 'Mã nhân viên', defaultValue: 'NV-2048', category: 'Cá nhân & Nhân sự' },
        Chuc_Vu: { label: 'Chức vụ', defaultValue: 'Chuyên viên Phân tích Nghiệp vụ', category: 'Cá nhân & Nhân sự' },
        Thoi_Gian_Tu_Den: { label: 'Thời gian xin nghỉ', defaultValue: 'Từ ngày 26/10/2026 đến hết ngày 28/10/2026 (03 ngày)', category: 'Thời gian & Địa điểm' },
        Ly_Do: { label: 'Lý do xin nghỉ', defaultValue: 'Giải quyết công việc gia đình tại quê nhà.', category: 'Tài chính & Văn bản' },
        Ben_Nhan: { label: 'Người nhận bàn giao', defaultValue: 'Lê Hoàng Nam (Phòng BA)', category: 'Đơn vị & Tổ chức' },
        So_Dien_Thoai: { label: 'Số điện thoại', defaultValue: '0912.345.678', category: 'Cá nhân & Nhân sự' },
        Dia_Diem: { label: 'Địa điểm lập đơn', defaultValue: 'TP. Hồ Chí Minh', category: 'Thời gian & Địa điểm' },
        Ngay: { label: 'Ngày', defaultValue: '23', category: 'Thời gian & Địa điểm' },
        Thang: { label: 'Tháng', defaultValue: '10', category: 'Thời gian & Địa điểm' },
        Nam: { label: 'Năm', defaultValue: '2026', category: 'Thời gian & Địa điểm' },
        Chu_Thich_Footnote: { label: 'Chú thích chân trang (Footnote)', defaultValue: 'Chú thích (1): Đơn xin nghỉ phép từ 03 ngày trở lên cần có xác nhận bàn giao công việc của Trưởng bộ phận.', category: 'Tài chính & Văn bản' },
      },
      custom_labels: {
        '0': 'Đầu trang (Headnote / Header)',
        '4': 'Kính gửi',
        '5': 'Họ tên nhân viên',
        '6': 'Mã nhân viên & Chức vụ (2 trường trên 1 dòng)',
        '7': 'Thời gian xin nghỉ (Có gắn Footnote #1)',
        '8': 'Lý do xin nghỉ',
        '9': 'Người nhận bàn giao',
        '10': 'Số điện thoại liên hệ',
        '12': 'Địa điểm, ngày tháng ký',
        '13': 'Chú thích chân trang (Footnote #1)',
      },
      base64_docx: leaveBase64,
      updated_at: '2026-10-23T06:02:00.000Z',
      total_paragraphs: SAMPLE_LEAVE_SPEC.length + 3,
      category: 'Hành chính - Nhân sự',
    },
    'Biên bản bàn giao thiết bị': {
      file_path: 'templates_store/bien_ban_ban_giao_thiet_bi.docx',
      selected_indices: ['0', '1', '2', '4', '5', '6', '7', '8', '11'],
      paragraph_templates: {
        '0': '«Ghi_Chu_Dau_Trang»',
        '1': '«Ten_Cong_Ty»',
        '2': 'Số: «So_Hieu_Van_Ban»',
        '4': 'Hôm nay, ngày «Ngay» tháng «Thang» năm «Nam», tại Văn phòng Công ty, chúng tôi gồm:',
        '5': 'Bên giao (Phòng IT - Hành chính): «Ben_Giao»',
        '6': 'Bên nhận (Nhân sự tiếp nhận): «Ben_Nhan»',
        '7': 'Thiết bị bàn giao: «Thiet_Bi_Tai_San»',
        '8': 'Tình trạng thiết bị khi bàn giao: «Tinh_Trang»',
        '11': '«Chu_Thich_Footnote»',
      },
      mailings_fields: {
        Ghi_Chu_Dau_Trang: { label: 'Tiêu đề đầu trang (Headnote / Header)', defaultValue: 'BIỂU MẪU QUẢN TRỊ THIẾT BỊ & TÀI SẢN — MÃ SỐ: BM-QTTS-03', category: 'Tài chính & Văn bản' },
        Ten_Cong_Ty: { label: 'Tên Công ty / Đơn vị', defaultValue: 'CÔNG TY CỔ PHẦN GIẢI PHÁP SỐ ĐÔNG DƯƠNG', category: 'Đơn vị & Tổ chức' },
        So_Hieu_Van_Ban: { label: 'Số hiệu văn bản', defaultValue: '42/2026/BBBG-HCNS', category: 'Tài chính & Văn bản' },
        Ngay: { label: 'Ngày', defaultValue: '23', category: 'Thời gian & Địa điểm' },
        Thang: { label: 'Tháng', defaultValue: '10', category: 'Thời gian & Địa điểm' },
        Nam: { label: 'Năm', defaultValue: '2026', category: 'Thời gian & Địa điểm' },
        Ben_Giao: { label: 'Bên giao', defaultValue: 'Ông Phạm Quốc Khánh - Chuyên viên Quản trị Hệ thống', category: 'Đơn vị & Tổ chức' },
        Ben_Nhan: { label: 'Bên nhận', defaultValue: 'Bà Vũ Ngọc Mai - Phòng Kinh doanh Doanh nghiệp', category: 'Đơn vị & Tổ chức' },
        Thiet_Bi_Tai_San: { label: 'Tên thiết bị / Tài sản', defaultValue: 'Máy tính xách tay MacBook Pro 14 inch M3 (S/N: C02X9876ABCD)', category: 'Tài chính & Văn bản' },
        Tinh_Trang: { label: 'Tình trạng thiết bị', defaultValue: 'Mới 100%, hoạt động tốt, kèm bộ sạc 70W nguyên hộp.', category: 'Tài chính & Văn bản' },
        Chu_Thich_Footnote: { label: 'Chú thích chân trang (Footnote)', defaultValue: 'Chú thích (1): Kiểm tra kỹ số Serial Number (S/N) và tem bảo hành trước khi ký nhận bàn giao.', category: 'Tài chính & Văn bản' },
      },
      custom_labels: {
        '0': 'Đầu trang (Headnote / Header)',
        '1': 'Tên Công ty',
        '2': 'Số hiệu biên bản',
        '4': 'Ngày tháng lập biên bản',
        '5': 'Thông tin Bên giao',
        '6': 'Thông tin Bên nhận',
        '7': 'Tên & Mã thiết bị bàn giao (Có gắn Footnote #1)',
        '8': 'Tình trạng thiết bị',
        '11': 'Chú thích chân trang (Footnote #1)',
      },
      base64_docx: handoverBase64,
      updated_at: '2026-10-23T06:05:00.000Z',
      total_paragraphs: SAMPLE_HANDOVER_SPEC.length + 3,
      category: 'Quản trị Thiết bị',
    },
  };

  // Preserve any custom user-created templates while upgrading the 3 starter templates
  const mergedConfig: TemplatesConfigMap = {
    ...synced,
    ...starterConfig,
  };

  saveTemplatesConfig(mergedConfig);
  localStorage.setItem(STORAGE_KEY_SEEDED, 'true');
  return mergedConfig;
}

export async function generateExtraTestUploadDocx(): Promise<{ fileName: string; blob: Blob }> {
  const spec: SampleParagraphSpec[] = [
    { text: 'CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM', align: 'center', bold: true, fontSizeHalfPt: 26 },
    { text: 'Độc lập - Tự do - Hạnh phúc', align: 'center', bold: true, fontSizeHalfPt: 26 },
    { text: 'GIẤY ĐỀ NGHỊ CẤP VĂN PHÒNG PHẨM', align: 'center', bold: true, fontSizeHalfPt: 30 },
    { text: 'Kính gửi: Phòng Hành chính - Quản trị', align: 'center', fontSizeHalfPt: 26 },
    { text: 'Họ và tên người đề nghị: Lê Quốc Bảo', align: 'left', fontSizeHalfPt: 26 },
    { text: 'Chức vụ: Trưởng nhóm Kỹ thuật | Phòng ban: Phòng Phát triển Sản phẩm', align: 'left', fontSizeHalfPt: 26 },
    { text: 'Nội dung đề nghị cấp: 05 ram giấy A4 Double A, 10 bút bi xanh Thiên Long, 03 sổ tay bìa cứng.', align: 'left', fontSizeHalfPt: 26 },
    { text: 'Lý do sử dụng: Phục vụ công tác in ấn hồ sơ nghiệm thu tháng 11/2026.', align: 'left', fontSizeHalfPt: 26 },
    { text: 'Thời gian cần cung cấp: Từ ngày 25/10/2026 đến ngày 30/10/2026', align: 'left', bold: true, fontSizeHalfPt: 26 },
    { text: 'Hà Nội, ngày 23 tháng 10 năm 2026', align: 'right', italic: true, fontSizeHalfPt: 26 },
    { text: 'Người đề nghị ký tên: Lê Quốc Bảo', align: 'right', bold: true, fontSizeHalfPt: 26 },
  ];
  const base64 = await generateSampleDocxBase64(spec);
  const binaryString = window.atob(base64);
  const bytes = new Uint8Array(binaryString.length);
  for (let i = 0; i < binaryString.length; i++) {
    bytes[i] = binaryString.charCodeAt(i);
  }
  const blob = new Blob([bytes], {
    type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  });
  return {
    fileName: 'giay_de_nghi_van_phong_pham_mau.docx',
    blob,
  };
}

export function autoDetectMailingsPositions(
  paragraphs: { index: number; text: string; section?: string }[],
  symbol: FieldSymbolPair = DEFAULT_FIELD_SYMBOL
): {
  selectedIndices: string[];
  paragraphTemplates: Record<string, string>;
  detectedFields: Record<string, { label: string; defaultValue: string; category?: MailingsFieldCategory }>;
} {
  const selectedIndices: string[] = [];
  const pTemplatesMap: Record<string, string> = {};
  const detectedFields: Record<
    string,
    { label: string; defaultValue: string; category?: MailingsFieldCategory }
  > = {};

  const tok = (k: string) => formatFieldToken(k, symbol);

  const registerField = (key: string, defaultVal: string) => {
    const def = getFieldDefByKey(key);
    if (!detectedFields[key]) {
      detectedFields[key] = {
        label: def.label,
        defaultValue: defaultVal.trim(),
        category: def.category,
      };
    }
  };

  paragraphs.forEach(({ index, text, section }) => {
    const trimmed = text.trim();
    if (!trimmed) return;
    const idxStr = String(index);

    const existingTokens = extractMergeFieldKeysFromText(trimmed, symbol);
    if (existingTokens.length > 0) {
      selectedIndices.push(idxStr);
      pTemplatesMap[idxStr] = text;
      existingTokens.forEach((k) => registerField(k, ''));
      return;
    }

    // Automatically select and map Header (Headnote) and Footnote / Endnote lines so they are always editable!
    if (section === 'header') {
      selectedIndices.push(idxStr);
      pTemplatesMap[idxStr] = tok('Ghi_Chu_Dau_Trang');
      registerField('Ghi_Chu_Dau_Trang', trimmed);
      return;
    }

    if (section === 'footnote' || section === 'endnote') {
      selectedIndices.push(idxStr);
      pTemplatesMap[idxStr] = tok('Chu_Thich_Footnote');
      registerField('Chu_Thich_Footnote', trimmed);
      return;
    }

    const dateRegex = /^([^,]+),\s*ngày\s+(\d{1,2}|\.+)\s+tháng\s+(\d{1,2}|\.+)\s+năm\s+(\d{4}|\.+)$/i;
    const dateMatch = trimmed.match(dateRegex);
    if (dateMatch) {
      selectedIndices.push(idxStr);
      pTemplatesMap[idxStr] = `${tok('Dia_Diem')}, ngày ${tok('Ngay')} tháng ${tok('Thang')} năm ${tok('Nam')}`;
      registerField('Dia_Diem', dateMatch[1]);
      registerField('Ngay', dateMatch[2]);
      registerField('Thang', dateMatch[3]);
      registerField('Nam', dateMatch[4]);
      return;
    }

    const inlineDateRegex = /(ngày\s+)(\d{1,2}|\.+)(\s+tháng\s+)(\d{1,2}|\.+)(\s+năm\s+)(\d{4}|\.+)/i;
    const inlineDateMatch = trimmed.match(inlineDateRegex);
    if (inlineDateMatch) {
      selectedIndices.push(idxStr);
      pTemplatesMap[idxStr] = text.replace(
        inlineDateRegex,
        `$1${tok('Ngay')}$3${tok('Thang')}$5${tok('Nam')}`
      );
      registerField('Ngay', inlineDateMatch[2]);
      registerField('Thang', inlineDateMatch[4]);
      registerField('Nam', inlineDateMatch[6]);
      return;
    }

    if (trimmed.includes('|') && trimmed.includes(':')) {
      const parts = text.split('|');
      let changed = false;
      const mappedParts = parts.map((segment) => {
        const colonPos = segment.indexOf(':');
        if (colonPos === -1) return segment;
        const prefix = segment.slice(0, colonPos + 1);
        const val = segment.slice(colonPos + 1).trim();
        const mappedKey = matchPrefixToMailingsKey(prefix);
        if (mappedKey) {
          changed = true;
          registerField(mappedKey, val);
          return `${prefix} ${tok(mappedKey)}`;
        }
        return segment;
      });
      if (changed) {
        selectedIndices.push(idxStr);
        pTemplatesMap[idxStr] = mappedParts.join(' | ');
        return;
      }
    }

    const colonIdx = text.indexOf(':');
    if (colonIdx > 0 && colonIdx < 65) {
      const prefix = text.slice(0, colonIdx + 1);
      const val = text.slice(colonIdx + 1).trim();
      const mappedKey = matchPrefixToMailingsKey(prefix);
      if (mappedKey) {
        selectedIndices.push(idxStr);
        pTemplatesMap[idxStr] = `${prefix} ${tok(mappedKey)}`;
        registerField(mappedKey, val);
      }
    }
  });

  return {
    selectedIndices,
    paragraphTemplates: pTemplatesMap,
    detectedFields,
  };
}

function matchPrefixToMailingsKey(prefixText: string): string | null {
  const lower = prefixText.toLowerCase();
  if (lower.includes('kính gửi')) return 'Kinh_Gui';
  if (
    lower.includes('họ và tên') ||
    lower.includes('tôi tên là') ||
    lower.includes('người đề nghị')
  ) {
    return 'Ho_Va_Ten';
  }
  if (lower.includes('mã nhân viên') || lower.includes('mã nv')) return 'Ma_Nhan_Vien';
  if (lower.includes('chức vụ')) return 'Chuc_Vu';
  if (lower.includes('bộ phận') || lower.includes('phòng ban') || lower.includes('đơn vị công tác')) {
    return 'Phong_Ban';
  }
  if (lower.includes('điện thoại') || lower.includes('sđt')) return 'So_Dien_Thoai';
  if (lower.includes('email')) return 'Email';
  if (lower.includes('cccd') || lower.includes('cmnd')) return 'CCCD_CMND';
  if (lower.includes('địa chỉ')) return 'Dia_Chi';
  if (lower.includes('bên giao')) return 'Ben_Giao';
  if (lower.includes('bên nhận') || lower.includes('bàn giao công việc')) return 'Ben_Nhan';
  if (lower.includes('thời gian')) return 'Thoi_Gian_Tu_Den';
  if (lower.includes('số tiền bằng chữ')) return 'So_Tien_Chu';
  if (lower.includes('số tiền')) return 'So_Tien_So';
  if (lower.includes('hình thức thanh toán') || lower.includes('tài khoản')) {
    return 'Hinh_Thuc_Thanh_Toan';
  }
  if (lower.includes('kèm theo')) return 'Kem_Theo';
  if (lower.includes('thiết bị') || lower.includes('tài sản')) return 'Thiet_Bi_Tai_San';
  if (lower.includes('tình trạng')) return 'Tinh_Trang';
  if (lower.includes('lý do') || lower.includes('mục đích')) return 'Ly_Do';
  if (lower.includes('nội dung') || lower.includes('danh mục')) return 'Noi_Dung';
  if (lower.trim() === 'số:') return 'So_Hieu_Van_Ban';
  return null;
}
