import React, { useRef, useState } from 'react';
import {
  Plus,
  RotateCcw,
  MousePointerClick,
  Check,
  Tag,
  X,
  Pencil,
  Trash2,
  ArrowRightLeft,
  Braces,
  Search,
} from 'lucide-react';
import { DocxParagraphItem } from '../utils/docxEngine';
import {
  DEFAULT_FIELD_SYMBOL,
  extractMergeFieldKeysFromText,
  FieldSymbolPair,
  formatFieldToken,
  getFieldDefByKey,
  MailingsFieldCategory,
  MailingsFieldDef,
  normalizeFieldKeyInput,
  PRESET_FIELD_SYMBOLS,
  removeFieldKeyFromTemplateString,
  replaceFieldKeyInTemplateString,
} from '../utils/templateStore';

interface MailingsPositionEditorProps {
  activeParagraph: DocxParagraphItem | null;
  currentTemplateString: string;
  customMailingsFields: MailingsFieldDef[];
  fieldSymbol: FieldSymbolPair;
  onUpdateFieldSymbol: (newSymbol: FieldSymbolPair) => void;
  onUpdateParagraphTemplate: (
    paragraphIndex: number,
    newTemplateStr: string,
    extractedDefaultVal?: { key: string; value: string }
  ) => void;
  onResetParagraphTemplate: (paragraphIndex: number) => void;
  onAddCustomField: (newField: MailingsFieldDef) => void;
  onRenameOrUpdateField: (oldKey: string, updatedField: MailingsFieldDef) => void;
  onDeleteField: (key: string) => void;
  onResetDefaultPalette: () => void;
}

const CATEGORIES: MailingsFieldCategory[] = [
  'Cá nhân & Nhân sự',
  'Đơn vị & Tổ chức',
  'Thời gian & Địa điểm',
  'Tài chính & Văn bản',
  'Trường tùy chỉnh',
];

export const MailingsPositionEditor: React.FC<MailingsPositionEditorProps> = ({
  activeParagraph,
  currentTemplateString,
  customMailingsFields,
  fieldSymbol = DEFAULT_FIELD_SYMBOL,
  onUpdateFieldSymbol,
  onUpdateParagraphTemplate,
  onResetParagraphTemplate,
  onAddCustomField,
  onRenameOrUpdateField,
  onDeleteField,
  onResetDefaultPalette,
}) => {
  const [selectedCategory, setSelectedCategory] = useState<MailingsFieldCategory | 'Tất cả'>('Tất cả');
  const [fieldSearchQuery, setFieldSearchQuery] = useState<string>('');
  const [selectionRange, setSelectionRange] = useState<{ start: number; end: number; selectedText: string }>({
    start: currentTemplateString.length,
    end: currentTemplateString.length,
    selectedText: '',
  });

  // Create or Edit Field Form state
  const [showFieldForm, setShowFieldForm] = useState<boolean>(false);
  const [editingOriginalKey, setEditingOriginalKey] = useState<string | null>(null);
  const [fieldFormLabel, setFieldFormLabel] = useState<string>('');
  const [fieldFormKey, setFieldFormKey] = useState<string>('');
  const [fieldFormPlaceholder, setFieldFormPlaceholder] = useState<string>('');
  const [fieldFormCategory, setFieldFormCategory] = useState<MailingsFieldCategory>('Trường tùy chỉnh');

  // Symbol customization state
  const [showSymbolCustomizer, setShowSymbolCustomizer] = useState<boolean>(false);
  const [customOpenInput, setCustomOpenInput] = useState<string>(fieldSymbol.openSymbol);
  const [customCloseInput, setCustomCloseInput] = useState<string>(fieldSymbol.closeSymbol);

  // Insertion mode state (includes replacing custom placeholder symbols like ...... or ___)
  const [insertMode, setInsertMode] = useState<
    'cursor_or_selection' | 'after_colon' | 'replace_symbol_pattern' | 'append_end'
  >('cursor_or_selection');
  const [targetPlaceholderSymbol, setTargetPlaceholderSymbol] = useState<string>('...');

  const textareaRef = useRef<HTMLTextAreaElement | null>(null);

  const allFields = customMailingsFields;
  const normalizedSearch = fieldSearchQuery.trim().toLowerCase();
  const normalizedSearchNoAccent = normalizeFieldKeyInput(fieldSearchQuery).toLowerCase();

  const visibleFields = allFields.filter((f) => {
    if (selectedCategory !== 'Tất cả' && f.category !== selectedCategory) {
      return false;
    }
    if (!normalizedSearch) return true;
    const tokenStr = formatFieldToken(f.key, fieldSymbol).toLowerCase();
    const labelNoAccent = normalizeFieldKeyInput(f.label).toLowerCase();
    return (
      f.label.toLowerCase().includes(normalizedSearch) ||
      f.key.toLowerCase().includes(normalizedSearch) ||
      tokenStr.includes(normalizedSearch) ||
      f.category.toLowerCase().includes(normalizedSearch) ||
      (normalizedSearchNoAccent.length > 0 &&
        (labelNoAccent.includes(normalizedSearchNoAccent) ||
          f.key.toLowerCase().includes(normalizedSearchNoAccent)))
    );
  });

  const handleTrackSelection = () => {
    const el = textareaRef.current;
    if (!el) return;
    const start = el.selectionStart ?? currentTemplateString.length;
    const end = el.selectionEnd ?? currentTemplateString.length;
    const selectedText = currentTemplateString.slice(start, end);
    setSelectionRange({ start, end, selectedText });
  };

  // Insert formatted token (e.g. «Field_Key» or {{Field_Key}}) at the chosen position in the active paragraph
  const handleInsertMergeField = (field: MailingsFieldDef) => {
    if (!activeParagraph) return;
    const token = formatFieldToken(field.key, fieldSymbol);
    const text = currentTemplateString;

    // Mode 1: Replace everything after the colon ":" in the line
    if (insertMode === 'after_colon') {
      const colonIdx = text.indexOf(':');
      if (colonIdx !== -1) {
        const prefix = text.slice(0, colonIdx + 1);
        const oldVal = text.slice(colonIdx + 1).trim();
        const updated = `${prefix} ${token}`;
        onUpdateParagraphTemplate(
          activeParagraph.index,
          updated,
          oldVal && !oldVal.includes(fieldSymbol.openSymbol)
            ? { key: field.key, value: oldVal }
            : undefined
        );
        return;
      }
    }

    // Mode 2: Replace target placeholder symbol (e.g. "......", "______", "…", or user-defined pattern)
    if (insertMode === 'replace_symbol_pattern') {
      const trimmedPattern = targetPlaceholderSymbol.trim();
      let updated = text;
      if (trimmedPattern === '...' || trimmedPattern === '___' || !trimmedPattern) {
        // Smart match first sequence of dots (.... / …) or underscores (____)
        const dotsOrUnderscoresRegex = /(\.{3,}|…+|__{2,})/;
        if (dotsOrUnderscoresRegex.test(text)) {
          updated = text.replace(dotsOrUnderscoresRegex, token);
          onUpdateParagraphTemplate(activeParagraph.index, updated);
          return;
        }
      } else if (text.includes(trimmedPattern)) {
        updated = text.replace(trimmedPattern, token);
        onUpdateParagraphTemplate(activeParagraph.index, updated);
        return;
      }
    }

    // Mode 3: Append to end of line
    if (insertMode === 'append_end') {
      const needSpace = text.length > 0 && !text.endsWith(' ');
      const updated = `${text}${needSpace ? ' ' : ''}${token}`;
      onUpdateParagraphTemplate(activeParagraph.index, updated);
      return;
    }

    // Mode 4 (Default): Insert at cursor or replace highlighted text in the position editor
    const el = textareaRef.current;
    const start = el ? el.selectionStart : selectionRange.start;
    const end = el ? el.selectionEnd : selectionRange.end;
    const safeStart = Math.max(0, Math.min(start, text.length));
    const safeEnd = Math.max(safeStart, Math.min(end, text.length));

    const replacedSubstring = text.slice(safeStart, safeEnd).trim();
    const updated = text.slice(0, safeStart) + token + text.slice(safeEnd);

    onUpdateParagraphTemplate(
      activeParagraph.index,
      updated,
      replacedSubstring && !replacedSubstring.includes(fieldSymbol.openSymbol)
        ? { key: field.key, value: replacedSubstring }
        : undefined
    );

    const nextCursorPos = safeStart + token.length;
    setTimeout(() => {
      if (textareaRef.current) {
        textareaRef.current.focus();
        textareaRef.current.setSelectionRange(nextCursorPos, nextCursorPos);
        setSelectionRange({ start: nextCursorPos, end: nextCursorPos, selectedText: '' });
      }
    }, 20);
  };

  // Open form to create a brand new field
  const handleOpenCreateFieldForm = () => {
    setEditingOriginalKey(null);
    setFieldFormLabel('');
    setFieldFormKey('');
    setFieldFormPlaceholder('');
    setFieldFormCategory('Trường tùy chỉnh');
    setShowFieldForm(true);
  };

  // Open form to edit / rename an existing field (both label and key)
  const handleOpenEditFieldForm = (field: MailingsFieldDef, e?: React.MouseEvent) => {
    if (e) e.stopPropagation();
    setEditingOriginalKey(field.key);
    setFieldFormLabel(field.label);
    setFieldFormKey(field.key);
    setFieldFormPlaceholder(field.placeholder || '');
    setFieldFormCategory(field.category || 'Trường tùy chỉnh');
    setShowFieldForm(true);
  };

  // Save new or renamed field
  const handleSaveFieldForm = () => {
    const label = fieldFormLabel.trim();
    if (!label) return;
    const generatedKey = normalizeFieldKeyInput(fieldFormKey || label);
    if (!generatedKey) return;

    const updatedField: MailingsFieldDef = {
      key: generatedKey,
      label,
      category: fieldFormCategory,
      placeholder: fieldFormPlaceholder.trim() || `Nhập ${label.toLowerCase()}...`,
    };

    if (editingOriginalKey) {
      onRenameOrUpdateField(editingOriginalKey, updatedField);
    } else {
      onAddCustomField(updatedField);
      if (activeParagraph) {
        handleInsertMergeField(updatedField);
      }
    }

    setShowFieldForm(false);
    setEditingOriginalKey(null);
    setFieldFormLabel('');
    setFieldFormKey('');
    setFieldFormPlaceholder('');
  };

  // Swap an existing field in the active line with another field («Old_Key» -> «New_Key»)
  const handleSwapFieldInActiveLine = (oldKey: string, newKey: string) => {
    if (!activeParagraph || !newKey || oldKey === newKey) return;
    const updated = replaceFieldKeyInTemplateString(
      currentTemplateString,
      oldKey,
      newKey,
      fieldSymbol
    );
    onUpdateParagraphTemplate(activeParagraph.index, updated);
  };

  // Remove a field token from the active line
  const handleRemoveFieldFromActiveLine = (targetKey: string) => {
    if (!activeParagraph) return;
    const updated = removeFieldKeyFromTemplateString(
      currentTemplateString,
      targetKey,
      fieldSymbol
    );
    if (
      updated === activeParagraph.text &&
      extractMergeFieldKeysFromText(updated, fieldSymbol).length === 0
    ) {
      onResetParagraphTemplate(activeParagraph.index);
    } else {
      onUpdateParagraphTemplate(activeParagraph.index, updated);
    }
  };

  const insertedKeysInActiveLine = activeParagraph
    ? extractMergeFieldKeysFromText(currentTemplateString, fieldSymbol)
    : [];

  return (
    <div className="bg-slate-50 border border-slate-200 rounded-lg p-4 space-y-4">
      {/* Top Header of Mailings Toolbar */}
      <div className="flex flex-wrap items-center justify-between gap-2 pb-3 border-b border-slate-200">
        <div>
          <h3 className="text-sm font-semibold text-slate-900 flex items-center gap-1.5">
            <Tag className="w-4 h-4 text-blue-700" />
            <span>Bộ Công Cụ Mailings Word — Chọn Vị Trí, Đổi Tên Trường & Ký Hiệu</span>
          </h3>
          <p className="text-xs text-slate-600 mt-0.5">
            Chèn trường dữ liệu, hoán đổi trường, đổi tên hiển thị / mã trường và tùy chỉnh ký hiệu bao quanh ({formatFieldToken('Ten_Truong', fieldSymbol)})
          </p>
        </div>

        <div className="flex flex-wrap items-center gap-2">
          <button
            type="button"
            onClick={() => {
              setCustomOpenInput(fieldSymbol.openSymbol);
              setCustomCloseInput(fieldSymbol.closeSymbol);
              setShowSymbolCustomizer((prev) => !prev);
            }}
            className={`inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold rounded-md border transition-colors whitespace-nowrap shrink-0 cursor-pointer ${
              showSymbolCustomizer
                ? 'bg-slate-900 text-white border-slate-900'
                : 'bg-white text-slate-800 border-slate-300 hover:bg-slate-100'
            }`}
          >
            <Braces className="w-3.5 h-3.5 text-blue-700" />
            <span>
              Ký hiệu trường: {fieldSymbol.openSymbol}...{fieldSymbol.closeSymbol}
            </span>
          </button>
        </div>
      </div>

      {/* Field Symbol / Delimiter Customizer Panel */}
      {showSymbolCustomizer && (
        <div className="p-3.5 bg-white border border-blue-300 rounded-lg space-y-3">
          <div className="flex items-center justify-between gap-2">
            <div>
              <h4 className="text-xs font-bold text-slate-900">
                Chọn hoặc tự nhập Ký hiệu bao quanh Trường dữ liệu
              </h4>
              <p className="text-[11px] text-slate-500">
                Khi đổi ký hiệu, toàn bộ các trường đã gắn trong biểu mẫu sẽ tự động chuyển sang ký hiệu mới
              </p>
            </div>
            <button
              type="button"
              onClick={() => setShowSymbolCustomizer(false)}
              className="text-slate-400 hover:text-slate-700 cursor-pointer"
            >
              <X className="w-4 h-4" />
            </button>
          </div>

          {/* Preset Symbol Buttons */}
          <div className="flex flex-wrap items-center gap-1.5">
            {PRESET_FIELD_SYMBOLS.map((preset) => {
              const isSelected =
                fieldSymbol.openSymbol === preset.openSymbol &&
                fieldSymbol.closeSymbol === preset.closeSymbol;
              return (
                <button
                  key={preset.id}
                  type="button"
                  onClick={() => {
                    setCustomOpenInput(preset.openSymbol);
                    setCustomCloseInput(preset.closeSymbol);
                    onUpdateFieldSymbol({
                      openSymbol: preset.openSymbol,
                      closeSymbol: preset.closeSymbol,
                    });
                  }}
                  className={`px-2.5 py-1.5 text-xs rounded-md border transition-colors cursor-pointer ${
                    isSelected
                      ? 'bg-blue-700 text-white border-blue-700 font-semibold'
                      : 'bg-slate-50 text-slate-700 border-slate-200 hover:border-blue-400'
                  }`}
                >
                  <span>{preset.name}</span>
                </button>
              );
            })}
          </div>

          {/* Custom Open / Close Symbol Inputs */}
          <div className="pt-2 border-t border-slate-100 flex flex-wrap items-end gap-3">
            <div>
              <label className="block text-[11px] font-medium text-slate-600 mb-1">
                Ký hiệu mở (Đầu trường):
              </label>
              <input
                type="text"
                value={customOpenInput}
                onChange={(e) => setCustomOpenInput(e.target.value)}
                placeholder="VD: « hoặc {{ hoặc @["
                className="w-28 px-2.5 py-1 text-xs font-mono-tabular bg-slate-50 border border-slate-300 rounded text-slate-900 focus:bg-white focus:outline-none focus:border-blue-600"
              />
            </div>
            <div>
              <label className="block text-[11px] font-medium text-slate-600 mb-1">
                Ký hiệu đóng (Cuối trường):
              </label>
              <input
                type="text"
                value={customCloseInput}
                onChange={(e) => setCustomCloseInput(e.target.value)}
                placeholder="VD: » hoặc }} hoặc ]"
                className="w-28 px-2.5 py-1 text-xs font-mono-tabular bg-slate-50 border border-slate-300 rounded text-slate-900 focus:bg-white focus:outline-none focus:border-blue-600"
              />
            </div>
            <button
              type="button"
              onClick={() => {
                const open = customOpenInput.trim();
                const close = customCloseInput.trim();
                if (!open || !close) return;
                onUpdateFieldSymbol({ openSymbol: open, closeSymbol: close });
              }}
              className="px-3 py-1.5 text-xs font-semibold text-white bg-blue-700 hover:bg-blue-800 rounded transition-colors cursor-pointer"
            >
              Áp dụng ký hiệu ({customOpenInput.trim() || '«'}Ho_Va_Ten{customCloseInput.trim() || '»'})
            </button>
          </div>
        </div>
      )}

      {/* Step 1: Active Line Position Box */}
      {!activeParagraph ? (
        <div className="p-4 bg-white border border-dashed border-slate-300 rounded-lg text-center space-y-1">
          <MousePointerClick className="w-5 h-5 text-blue-700 mx-auto" />
          <p className="text-xs font-semibold text-slate-800">
            Hãy nhấp chọn một dòng bất kỳ ở danh sách bên dưới hoặc trên trang A4 bên phải
          </p>
          <p className="text-[11px] text-slate-500">
            Sau khi chọn dòng, bạn có thể đặt con trỏ tại bất kỳ vị trí nào trong câu để chèn hoặc thay đổi trường {formatFieldToken('Ho_Va_Ten', fieldSymbol)}, {formatFieldToken('Ngay', fieldSymbol)}...
          </p>
        </div>
      ) : (
        <div className="p-3.5 bg-white border border-blue-600/60 rounded-lg space-y-3">
          <div className="flex flex-wrap items-center justify-between gap-2">
            <div className="flex flex-wrap items-center gap-2 text-xs">
              <span className="font-mono-tabular font-semibold text-blue-800">
                Đang chỉnh vị trí tại [Dòng {String(activeParagraph.index).padStart(2, '0')}]
              </span>
              {activeParagraph.sectionLabel && (
                <>
                  <span aria-hidden="true" className="text-slate-300">·</span>
                  <span className="font-semibold text-slate-800">
                    {activeParagraph.tablePosition
                      ? `${activeParagraph.sectionLabel} — ${activeParagraph.tablePosition}`
                      : activeParagraph.sectionLabel}
                  </span>
                </>
              )}
              <span aria-hidden="true" className="text-slate-300">·</span>
              <span className="text-slate-600">
                {selectionRange.selectedText
                  ? `Đang bôi đen: "${selectionRange.selectedText.slice(0, 35)}${selectionRange.selectedText.length > 35 ? '...' : ''}" (Nhấp trường để thay thế)`
                  : `Vị trí con trỏ: ký tự thứ ${selectionRange.start}`}
              </span>
            </div>

            <div className="flex items-center gap-2">
              {currentTemplateString !== activeParagraph.text && (
                <button
                  type="button"
                  onClick={() => onResetParagraphTemplate(activeParagraph.index)}
                  className="inline-flex items-center gap-1 text-xs font-medium text-slate-600 hover:text-slate-900 whitespace-nowrap shrink-0 cursor-pointer"
                >
                  <RotateCcw className="w-3 h-3" />
                  <span>Khôi phục câu gốc</span>
                </button>
              )}
              <button
                type="button"
                onClick={() => onResetParagraphTemplate(activeParagraph.index)}
                title="Bỏ chọn và xóa cấu hình trường tại dòng này"
                className="inline-flex items-center gap-1 px-2 py-1 text-xs font-medium text-red-600 bg-red-50 hover:bg-red-100 rounded border border-red-200 whitespace-nowrap shrink-0 cursor-pointer"
              >
                <Trash2 className="w-3 h-3" />
                <span>Bỏ chọn dòng này</span>
              </button>
            </div>
          </div>

          {/* Position Mode Selector */}
          <div className="flex flex-wrap items-center gap-2 text-xs">
            <span className="text-slate-500 font-medium">Cách đặt vị trí khi nhấp trường:</span>
            <div className="flex flex-wrap items-center gap-1 p-0.5 bg-slate-100 rounded-md">
              <button
                type="button"
                onClick={() => setInsertMode('cursor_or_selection')}
                className={`px-2.5 py-1 text-[11px] font-medium rounded transition-colors whitespace-nowrap shrink-0 cursor-pointer ${
                  insertMode === 'cursor_or_selection'
                    ? 'bg-white text-blue-800 shadow-2xs font-semibold'
                    : 'text-slate-600 hover:text-slate-900'
                }`}
              >
                Tại con trỏ / vùng bôi đen
              </button>
              {currentTemplateString.includes(':') && (
                <button
                  type="button"
                  onClick={() => setInsertMode('after_colon')}
                  className={`px-2.5 py-1 text-[11px] font-medium rounded transition-colors whitespace-nowrap shrink-0 cursor-pointer ${
                    insertMode === 'after_colon'
                      ? 'bg-white text-blue-800 shadow-2xs font-semibold'
                      : 'text-slate-600 hover:text-slate-900'
                  }`}
                >
                  Thay nội dung sau dấu &ldquo;:&rdquo;
                </button>
              )}
              <button
                type="button"
                onClick={() => setInsertMode('replace_symbol_pattern')}
                className={`px-2.5 py-1 text-[11px] font-medium rounded transition-colors whitespace-nowrap shrink-0 cursor-pointer ${
                  insertMode === 'replace_symbol_pattern'
                    ? 'bg-white text-blue-800 shadow-2xs font-semibold'
                    : 'text-slate-600 hover:text-slate-900'
                }`}
              >
                Thay ký hiệu cần (...... / ___)
              </button>
              <button
                type="button"
                onClick={() => setInsertMode('append_end')}
                className={`px-2.5 py-1 text-[11px] font-medium rounded transition-colors whitespace-nowrap shrink-0 cursor-pointer ${
                  insertMode === 'append_end'
                    ? 'bg-white text-blue-800 shadow-2xs font-semibold'
                    : 'text-slate-600 hover:text-slate-900'
                }`}
              >
                Chèn nối vào cuối dòng
              </button>
            </div>

            {insertMode === 'replace_symbol_pattern' && (
              <div className="flex items-center gap-1.5">
                <span className="text-[11px] text-slate-600">Ký hiệu cần thay:</span>
                <input
                  type="text"
                  value={targetPlaceholderSymbol}
                  onChange={(e) => setTargetPlaceholderSymbol(e.target.value)}
                  placeholder="... hoặc ___"
                  className="w-24 px-2 py-0.5 text-xs font-mono-tabular bg-slate-50 border border-slate-300 rounded text-slate-900 focus:bg-white focus:outline-none focus:border-blue-600"
                />
              </div>
            )}
          </div>

          {/* Interactive Line Textbox where user places cursor or highlights text */}
          <div>
            <textarea
              ref={textareaRef}
              rows={2}
              value={currentTemplateString}
              onClick={handleTrackSelection}
              onKeyUp={handleTrackSelection}
              onSelect={handleTrackSelection}
              onChange={(e) => {
                onUpdateParagraphTemplate(activeParagraph.index, e.target.value);
                handleTrackSelection();
              }}
              className="w-full px-3 py-2 text-sm font-medium bg-slate-50 border border-slate-300 rounded-md text-slate-900 focus:bg-white focus:outline-none focus:border-blue-600"
              placeholder="Nội dung dòng văn bản..."
            />
          </div>

          {/* Active Line Field Swapper & Renamer: Change or Rename fields already placed in this line */}
          {insertedKeysInActiveLine.length > 0 && (
            <div className="pt-2 border-t border-slate-100 space-y-1.5">
              <span className="block text-[11px] font-semibold text-slate-700">
                Trường đang gắn tại dòng này (Có thể đổi sang trường khác, đổi tên hoặc xóa):
              </span>
              <div className="flex flex-wrap items-center gap-2">
                {insertedKeysInActiveLine.map((k) => {
                  const fDef = getFieldDefByKey(k, allFields);
                  const isInPalette = allFields.some((opt) => opt.key === k);
                  return (
                    <div
                      key={k}
                      className="inline-flex items-center gap-1.5 px-2.5 py-1 bg-blue-50/70 border border-blue-200 rounded-md text-xs"
                    >
                      <span className="font-semibold text-slate-900">{fDef.label}</span>
                      <span className="font-mono-tabular text-[11px] text-blue-700 font-medium">
                        {formatFieldToken(k, fieldSymbol)}
                      </span>

                      {/* Dropdown to swap this field with another field */}
                      <div className="relative inline-flex items-center">
                        <ArrowRightLeft className="w-3 h-3 text-slate-500 mr-0.5" />
                        <select
                          value={k}
                          onChange={(e) => handleSwapFieldInActiveLine(k, e.target.value)}
                          title="Thay đổi sang trường khác tại vị trí này"
                          className="w-[250px] text-[11px] font-medium text-blue-800 bg-white border border-blue-200 rounded px-1.5 py-0.5 cursor-pointer focus:outline-none focus:border-blue-600"
                        >
                          {!isInPalette && (
                            <option value={k}>
                              Đang chọn: {fDef.label} ({formatFieldToken(k, fieldSymbol)})
                            </option>
                          )}
                          {allFields.map((opt) => (
                            <option key={opt.key} value={opt.key}>
                              Đổi thành: {opt.label} ({formatFieldToken(opt.key, fieldSymbol)})
                            </option>
                          ))}
                        </select>
                      </div>

                      {/* Button to rename this field */}
                      <button
                        type="button"
                        onClick={(e) => handleOpenEditFieldForm(fDef, e)}
                        title="Đổi tên hiển thị & mã ký hiệu trường này"
                        className="p-1 text-slate-600 hover:text-blue-700 rounded hover:bg-blue-100/60 cursor-pointer"
                      >
                        <Pencil className="w-3 h-3" />
                      </button>

                      {/* Button to remove this field from line */}
                      <button
                        type="button"
                        onClick={() => handleRemoveFieldFromActiveLine(k)}
                        title="Xóa trường này khỏi dòng"
                        className="inline-flex items-center gap-0.5 px-1.5 py-0.5 text-[11px] font-semibold text-red-600 hover:text-red-700 bg-red-50 hover:bg-red-100 border border-red-200 rounded cursor-pointer"
                      >
                        <X className="w-3 h-3" />
                        <span>Xóa</span>
                      </button>
                    </div>
                  );
                })}
              </div>
            </div>
          )}
        </div>
      )}

      {/* Step 2: Word Mailings Fields Palette (Insert, Edit/Rename, or Create Merge Field) */}
      <div className="space-y-2.5">
        <div className="flex flex-wrap items-center justify-between gap-2">
          <div className="flex flex-wrap items-center gap-1">
            <button
              type="button"
              onClick={() => setSelectedCategory('Tất cả')}
              className={`px-2.5 py-1 text-xs font-medium rounded-md transition-colors whitespace-nowrap shrink-0 cursor-pointer ${
                selectedCategory === 'Tất cả'
                  ? 'bg-slate-900 text-white'
                  : 'bg-white text-slate-600 border border-slate-200 hover:text-slate-900'
              }`}
            >
              Tất cả trường ({allFields.length})
            </button>
            {CATEGORIES.map((cat) => (
              <button
                key={cat}
                type="button"
                onClick={() => setSelectedCategory(cat)}
                className={`px-2.5 py-1 text-xs font-medium rounded-md transition-colors whitespace-nowrap shrink-0 cursor-pointer ${
                  selectedCategory === cat
                    ? 'bg-slate-900 text-white'
                    : 'bg-white text-slate-600 border border-slate-200 hover:text-slate-900'
                }`}
              >
                {cat}
              </button>
            ))}
          </div>

          <div className="flex items-center gap-1.5">
            <button
              type="button"
              onClick={onResetDefaultPalette}
              title="Khôi phục lại danh sách các trường mặc định ban đầu"
              className="inline-flex items-center gap-1 px-2.5 py-1 text-xs font-medium text-slate-600 bg-white border border-slate-200 hover:text-slate-900 rounded-md whitespace-nowrap shrink-0 cursor-pointer"
            >
              <RotateCcw className="w-3 h-3" />
              <span>Mặc định</span>
            </button>
            <button
              type="button"
              onClick={() => {
                if (showFieldForm && !editingOriginalKey) {
                  setShowFieldForm(false);
                } else {
                  handleOpenCreateFieldForm();
                }
              }}
              className="inline-flex items-center gap-1 px-2.5 py-1 text-xs font-semibold text-blue-700 bg-white border border-blue-300 hover:bg-blue-50 rounded-md whitespace-nowrap shrink-0 cursor-pointer"
            >
              {showFieldForm && !editingOriginalKey ? (
                <X className="w-3 h-3" />
              ) : (
                <Plus className="w-3 h-3" />
              )}
              <span>Tạo trường mới</span>
            </button>
          </div>
        </div>

        {/* Form to Create a New Field OR Rename / Change an Existing Field */}
        {showFieldForm && (
          <div className="p-3.5 bg-white border border-blue-400 rounded-lg space-y-3">
            <div className="flex items-center justify-between gap-2">
              <h4 className="text-xs font-bold text-slate-900 flex items-center gap-1.5">
                <Pencil className="w-3.5 h-3.5 text-blue-700" />
                <span>
                  {editingOriginalKey
                    ? `Đổi tên & Mã ký hiệu cho trường ${formatFieldToken(editingOriginalKey, fieldSymbol)}`
                    : 'Tạo trường dữ liệu Mailings mới'}
                </span>
              </h4>
              <button
                type="button"
                onClick={() => {
                  setShowFieldForm(false);
                  setEditingOriginalKey(null);
                }}
                className="text-slate-400 hover:text-slate-700 cursor-pointer"
              >
                <X className="w-4 h-4" />
              </button>
            </div>

            <div className="grid grid-cols-1 sm:grid-cols-12 gap-2.5 items-end">
              <div className="sm:col-span-4">
                <label className="block text-[11px] font-medium text-slate-700 mb-1">
                  Tên hiển thị của trường (VD: Họ tên khách hàng):
                </label>
                <input
                  type="text"
                  value={fieldFormLabel}
                  onChange={(e) => {
                    const val = e.target.value;
                    setFieldFormLabel(val);
                    if (!editingOriginalKey) {
                      setFieldFormKey(normalizeFieldKeyInput(val));
                    }
                  }}
                  placeholder="Nhập tên trường..."
                  className="w-full px-2.5 py-1.5 text-xs bg-slate-50 border border-slate-300 rounded text-slate-900 focus:bg-white focus:outline-none focus:border-blue-600"
                />
              </div>
              <div className="sm:col-span-4">
                <label className="block text-[11px] font-medium text-slate-700 mb-1">
                  Mã ký hiệu {fieldSymbol.openSymbol}Key{fieldSymbol.closeSymbol}:
                </label>
                <input
                  type="text"
                  value={fieldFormKey}
                  onChange={(e) => setFieldFormKey(e.target.value)}
                  placeholder="VD: Ho_Ten_Khach_Hang"
                  className="w-full px-2.5 py-1.5 text-xs font-mono-tabular bg-slate-50 border border-slate-300 rounded text-slate-900 focus:bg-white focus:outline-none focus:border-blue-600"
                />
              </div>
              <div className="sm:col-span-4">
                <label className="block text-[11px] font-medium text-slate-700 mb-1">
                  Phân nhóm trường:
                </label>
                <select
                  value={fieldFormCategory}
                  onChange={(e) => setFieldFormCategory(e.target.value as MailingsFieldCategory)}
                  className="w-full px-2.5 py-1.5 text-xs bg-slate-50 border border-slate-300 rounded text-slate-900 focus:bg-white focus:outline-none focus:border-blue-600"
                >
                  {CATEGORIES.map((cat) => (
                    <option key={cat} value={cat}>
                      {cat}
                    </option>
                  ))}
                </select>
              </div>
              <div className="sm:col-span-8">
                <label className="block text-[11px] font-medium text-slate-700 mb-1">
                  Giá trị gợi ý mặc định:
                </label>
                <input
                  type="text"
                  value={fieldFormPlaceholder}
                  onChange={(e) => setFieldFormPlaceholder(e.target.value)}
                  placeholder="Nhập nội dung ví dụ..."
                  className="w-full px-2.5 py-1.5 text-xs bg-slate-50 border border-slate-300 rounded text-slate-900 focus:bg-white focus:outline-none focus:border-blue-600"
                />
              </div>
              <div className="sm:col-span-4 flex items-center gap-1.5">
                <button
                  type="button"
                  onClick={handleSaveFieldForm}
                  className="flex-1 flex items-center justify-center gap-1 px-3 py-1.5 text-xs font-semibold text-white bg-blue-700 hover:bg-blue-800 rounded transition-colors whitespace-nowrap shrink-0 cursor-pointer"
                >
                  <Check className="w-3.5 h-3.5" />
                  <span>{editingOriginalKey ? 'Lưu đổi tên trường' : 'Thêm & Chèn'}</span>
                </button>
              </div>
            </div>
          </div>
        )}

        {/* Search Input Above Field List */}
        <div className="relative">
          <Search className="w-3.5 h-3.5 text-slate-400 absolute left-3 top-1/2 -translate-y-1/2 pointer-events-none" />
          <input
            type="text"
            value={fieldSearchQuery}
            onChange={(e) => setFieldSearchQuery(e.target.value)}
            placeholder={`Tìm nhanh trường theo tên hoặc mã ${formatFieldToken('Key', fieldSymbol)} (VD: Họ và tên, Ngày sinh, Số tiền)...`}
            className="w-full pl-8 pr-20 py-1.5 text-xs bg-white border border-slate-300 rounded-md text-slate-900 focus:outline-none focus:border-blue-600 transition-colors"
          />
          {fieldSearchQuery && (
            <div className="absolute right-2 top-1/2 -translate-y-1/2 flex items-center gap-1.5">
              <span className="text-[10px] font-mono-tabular text-slate-500">
                {visibleFields.length} kết quả
              </span>
              <button
                type="button"
                onClick={() => setFieldSearchQuery('')}
                title="Xóa từ khóa tìm kiếm"
                className="p-0.5 text-slate-400 hover:text-slate-700 rounded cursor-pointer"
              >
                <X className="w-3.5 h-3.5" />
              </button>
            </div>
          )}
        </div>

        {/* Grid of Clickable Mailings Merge Fields with inline Edit/Rename and Delete buttons */}
        {visibleFields.length === 0 ? (
          <div className="p-4 bg-white border border-dashed border-slate-300 rounded-md text-center space-y-2">
            <p className="text-xs text-slate-600">
              Không tìm thấy trường nào khớp với &ldquo;<strong>{fieldSearchQuery}</strong>&rdquo;
              {selectedCategory !== 'Tất cả' ? ` trong nhóm "${selectedCategory}"` : ''}.
            </p>
            <div className="flex flex-wrap items-center justify-center gap-2">
              {selectedCategory !== 'Tất cả' && (
                <button
                  type="button"
                  onClick={() => setSelectedCategory('Tất cả')}
                  className="px-2.5 py-1 text-xs font-medium text-slate-700 bg-slate-100 hover:bg-slate-200 rounded cursor-pointer"
                >
                  Tìm trong tất cả nhóm
                </button>
              )}
              {fieldSearchQuery.trim() && (
                <button
                  type="button"
                  onClick={() => {
                    const label = fieldSearchQuery.trim();
                    setEditingOriginalKey(null);
                    setFieldFormLabel(label);
                    setFieldFormKey(normalizeFieldKeyInput(label));
                    setFieldFormPlaceholder(`Nhập ${label.toLowerCase()}...`);
                    setFieldFormCategory(
                      selectedCategory === 'Tất cả' ? 'Trường tùy chỉnh' : selectedCategory
                    );
                    setShowFieldForm(true);
                  }}
                  className="inline-flex items-center gap-1 px-2.5 py-1 text-xs font-semibold text-white bg-blue-700 hover:bg-blue-800 rounded cursor-pointer"
                >
                  <Plus className="w-3 h-3" />
                  <span>Tạo nhanh trường &ldquo;{fieldSearchQuery.trim()}&rdquo;</span>
                </button>
              )}
            </div>
          </div>
        ) : (
          <div className="grid grid-cols-2 sm:grid-cols-3 gap-1.5 max-h-[225px] overflow-y-auto pr-1">
            {visibleFields.map((field) => {
              const isAlreadyInLine = insertedKeysInActiveLine.includes(field.key);
              const formattedToken = formatFieldToken(field.key, fieldSymbol);
              return (
                <div
                key={field.key}
                onClick={() => {
                  if (activeParagraph) {
                    handleInsertMergeField(field);
                  }
                }}
                title={
                  activeParagraph
                    ? `Nhấp để chèn ${formattedToken} (${field.label}) vào Dòng ${activeParagraph.index}`
                    : 'Hãy chọn một dòng văn bản trước để chèn trường này (hoặc bấm biểu tượng cây bút để đổi tên trường)'
                }
                className={`group relative flex items-center justify-between gap-1 px-2.5 py-1.5 rounded border text-left transition-colors ${
                  !activeParagraph
                    ? 'bg-white border-slate-200 hover:border-slate-300'
                    : isAlreadyInLine
                    ? 'bg-blue-50/80 border-blue-400 hover:bg-blue-100 cursor-pointer'
                    : 'bg-white border-slate-200 hover:border-blue-500 hover:bg-blue-50/40 cursor-pointer'
                }`}
              >
                <div className="min-w-0 flex-1">
                  <span className="block text-xs font-medium text-slate-900 truncate">
                    {field.label}
                  </span>
                  <span className="block text-[11px] font-mono-tabular text-blue-700 truncate">
                    {formattedToken}
                  </span>
                </div>

                {/* Action buttons on each field card: Edit/Rename & Delete */}
                <div
                  className="flex items-center gap-0.5 opacity-70 group-hover:opacity-100 shrink-0"
                  onClick={(e) => e.stopPropagation()}
                >
                  <button
                    type="button"
                    onClick={(e) => handleOpenEditFieldForm(field, e)}
                    title={`Đổi tên hoặc đổi mã trường ${formattedToken}`}
                    className="p-1 text-slate-500 hover:text-blue-700 hover:bg-blue-100/70 rounded cursor-pointer"
                  >
                    <Pencil className="w-3 h-3" />
                  </button>
                  <button
                    type="button"
                    onClick={() => onDeleteField(field.key)}
                    title={`Xóa trường ${formattedToken} khỏi danh sách`}
                    className="p-1 text-slate-400 hover:text-red-600 hover:bg-red-50 rounded cursor-pointer"
                  >
                    <Trash2 className="w-3 h-3" />
                  </button>
                </div>
              </div>
            );
          })}
          </div>
        )}
      </div>
    </div>
  );
};
