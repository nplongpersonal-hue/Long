import React, { useEffect, useMemo, useRef, useState } from 'react';
import { Eye, Code2, FileCheck2 } from 'lucide-react';
import { renderAsync } from 'docx-preview';
import {
  DocxParagraphItem,
  DocxRunSpan,
  exportFilledDocxBlob,
} from '../utils/docxEngine';
import {
  buildUniversalFieldTokenRegex,
  DEFAULT_FIELD_SYMBOL,
  FieldSymbolPair,
  formatFieldToken,
  getFieldDefByKey,
  hasAnyMergeFieldToken,
  MailingsFieldDef,
  matchSingleFieldToken,
  resolveParagraphTemplate,
} from '../utils/templateStore';

interface WordPreviewSheetProps {
  templateName: string;
  base64Docx?: string;
  paragraphs: DocxParagraphItem[];
  selectedIndices: string[];
  paragraphTemplates?: Record<string, string>;
  fieldValues?: Record<string, string>;
  replacements?: Record<number, string>;
  customLabels?: Record<string, string>;
  customMailingsFields?: MailingsFieldDef[];
  fieldSymbol?: FieldSymbolPair;
  activeIndex?: number | null;
  activeFieldKey?: string | null;
  mode: 'fill' | 'configure';
  onSelectParagraph?: (index: number) => void;
  onSelectFieldKey?: (fieldKey: string, paragraphIndex: number) => void;
  showLineNumbers?: boolean;
}

type PreviewDisplayMode = 'exact_docx' | 'merged_interactive' | 'tokens_interactive';

interface TableGroupBlock {
  kind: 'table';
  tableId: number;
  hasBorders: boolean;
  rows: Map<number, Map<number, { colSpan: number; paragraphs: DocxParagraphItem[] }>>;
}

interface SingleParagraphBlock {
  kind: 'paragraph';
  paragraph: DocxParagraphItem;
}

type BodyRenderBlock = TableGroupBlock | SingleParagraphBlock;

/**
 * Groups consecutive body paragraphs that belong to the same <w:tbl> into an actual table grid
 * so tables display with real rows and columns instead of flattened lines.
 */
function groupBodyParagraphsAndTables(bodies: DocxParagraphItem[]): BodyRenderBlock[] {
  const blocks: BodyRenderBlock[] = [];
  let currentTable: TableGroupBlock | null = null;

  bodies.forEach((p) => {
    if (p.isInTable && p.tableId !== undefined) {
      const rIdx = p.tableRowIndex ?? 1;
      const cIdx = p.tableColIndex ?? 1;
      const cSpan = p.tableColSpan ?? 1;

      if (!currentTable || currentTable.tableId !== p.tableId) {
        currentTable = {
          kind: 'table',
          tableId: p.tableId,
          hasBorders: p.tableHasBorders !== false,
          rows: new Map(),
        };
        blocks.push(currentTable);
      }

      if (!currentTable.rows.has(rIdx)) {
        currentTable.rows.set(rIdx, new Map());
      }
      const rowMap = currentTable.rows.get(rIdx)!;
      if (!rowMap.has(cIdx)) {
        rowMap.set(cIdx, { colSpan: cSpan, paragraphs: [] });
      }
      rowMap.get(cIdx)!.paragraphs.push(p);
    } else {
      currentTable = null;
      blocks.push({ kind: 'paragraph', paragraph: p });
    }
  });

  return blocks;
}

/**
 * Renders text preserving inline bold/italic/underline/fontSize/color from original Word runs,
 * even after field values have been replaced!
 */
function renderStyledTextFromOriginalRuns(
  originalText: string,
  currentText: string,
  runs?: DocxRunSpan[]
): React.ReactNode {
  const renderWithTabs = (str: string) => {
    if (!str.includes('\t')) return str;
    const parts = str.split('\t');
    return parts.map((pt, idx) => (
      <React.Fragment key={idx}>
        {idx > 0 && <span className="inline-block w-8" aria-hidden="true">&emsp;</span>}
        {pt}
      </React.Fragment>
    ));
  };

  if (!runs || runs.length === 0) {
    return renderWithTabs(currentText);
  }

  // If text has not changed, render each original run with its exact Word style
  if (currentText === originalText) {
    return runs.map((run, idx) => (
      <span
        key={idx}
        className={`${run.bold ? 'font-bold' : 'font-normal'} ${run.italic ? 'italic' : ''} ${
          run.underline ? 'underline underline-offset-2' : ''
        }`}
        style={{
          fontSize: run.fontSizePt ? `${run.fontSizePt}pt` : undefined,
          color: run.color || undefined,
        }}
      >
        {renderWithTabs(run.text)}
      </span>
    ));
  }

  // If text was modified (e.g., "Họ và tên: Nguyễn Văn A" -> "Họ và tên: Trần Thị B"),
  // preserve the style of matching prefixes/labels and apply value style to the replaced parts
  const firstRun = runs[0];
  if (
    runs.length >= 2 &&
    firstRun.text.trim().length > 0 &&
    currentText.startsWith(firstRun.text)
  ) {
    const restText = currentText.slice(firstRun.text.length);
    const secondRun = runs[1];
    return (
      <>
        <span
          className={`${firstRun.bold ? 'font-bold' : 'font-normal'} ${
            firstRun.italic ? 'italic' : ''
          } ${firstRun.underline ? 'underline underline-offset-2' : ''}`}
          style={{
            fontSize: firstRun.fontSizePt ? `${firstRun.fontSizePt}pt` : undefined,
            color: firstRun.color || undefined,
          }}
        >
          {renderWithTabs(firstRun.text)}
        </span>
        <span
          className={`${secondRun.bold ? 'font-bold' : 'font-normal'} ${
            secondRun.italic ? 'italic' : ''
          } ${secondRun.underline ? 'underline underline-offset-2' : ''}`}
          style={{
            fontSize: secondRun.fontSizePt ? `${secondRun.fontSizePt}pt` : undefined,
            color: secondRun.color || undefined,
          }}
        >
          {renderWithTabs(restText)}
        </span>
      </>
    );
  }

  // If line has "Label: Value" structure where original had bold label before colon
  const colonIdx = currentText.indexOf(':');
  const origHasBoldFirstRun = Boolean(firstRun?.bold && runs.some((r) => !r.bold));
  if (origHasBoldFirstRun && colonIdx > 0 && colonIdx < 50) {
    const prefix = currentText.slice(0, colonIdx + 1);
    const suffix = currentText.slice(colonIdx + 1);
    const valueRun = runs.find((r) => !r.bold) || firstRun;
    return (
      <>
        <span
          className={`font-bold ${firstRun.italic ? 'italic' : ''} ${
            firstRun.underline ? 'underline underline-offset-2' : ''
          }`}
          style={{
            fontSize: firstRun.fontSizePt ? `${firstRun.fontSizePt}pt` : undefined,
          }}
        >
          {renderWithTabs(prefix)}
        </span>
        <span
          className={`font-normal ${valueRun.italic ? 'italic' : ''} ${
            valueRun.underline ? 'underline underline-offset-2' : ''
          }`}
          style={{
            fontSize: valueRun.fontSizePt ? `${valueRun.fontSizePt}pt` : undefined,
          }}
        >
          {renderWithTabs(suffix)}
        </span>
      </>
    );
  }

  return renderWithTabs(currentText);
}

const EMPTY_STRING_MAP: Record<string, string> = {};
const EMPTY_NUMBER_MAP: Record<number, string> = {};
const EMPTY_FIELDS_ARRAY: MailingsFieldDef[] = [];

export const WordPreviewSheet: React.FC<WordPreviewSheetProps> = ({
  templateName,
  base64Docx,
  paragraphs,
  selectedIndices,
  paragraphTemplates = EMPTY_STRING_MAP,
  fieldValues = EMPTY_STRING_MAP,
  replacements = EMPTY_NUMBER_MAP,
  customLabels = EMPTY_STRING_MAP,
  customMailingsFields = EMPTY_FIELDS_ARRAY,
  fieldSymbol = DEFAULT_FIELD_SYMBOL,
  activeIndex = null,
  activeFieldKey = null,
  mode,
  onSelectParagraph,
  onSelectFieldKey,
  showLineNumbers = true,
}) => {
  const [displayMode, setDisplayMode] = useState<PreviewDisplayMode>(
    mode === 'fill' ? 'merged_interactive' : 'tokens_interactive'
  );
  const [isRenderingDocx, setIsRenderingDocx] = useState<boolean>(false);

  const docxHostRef = useRef<HTMLDivElement | null>(null);
  const paragraphsRef = useRef<DocxParagraphItem[]>(paragraphs);
  paragraphsRef.current = paragraphs;
  const activeIndexRef = useRef<number | null>(activeIndex);
  activeIndexRef.current = activeIndex;
  const lastRenderedSignatureRef = useRef<string>('');

  const selectedSet = useMemo(
    () => new Set(selectedIndices.map((s) => Number(s))),
    [selectedIndices]
  );

  // Compute the exact replacements map that will be written to the exported .docx file
  const effectiveReplacements = useMemo(() => {
    const map: Record<number, string> = {};
    paragraphs.forEach((p) => {
      const idxStr = String(p.index);
      const pTpl = paragraphTemplates[idxStr];
      if (mode === 'fill' && replacements[p.index] !== undefined) {
        map[p.index] = replacements[p.index];
      } else if (pTpl) {
        map[p.index] = resolveParagraphTemplate(pTpl, fieldValues, mode === 'configure', fieldSymbol);
      }
    });
    return map;
  }, [paragraphs, paragraphTemplates, fieldValues, replacements, mode, fieldSymbol]);

  const effectiveReplacementsRef = useRef<Record<number, string>>(effectiveReplacements);
  effectiveReplacementsRef.current = effectiveReplacements;

  // Primitive string signature so the docx-preview effect ONLY runs when actual content changes
  const effectiveReplacementsSignature = useMemo(
    () => JSON.stringify(effectiveReplacements),
    [effectiveReplacements]
  );

  // Reset cached signature when switching away from exact_docx so switching back renders fresh
  useEffect(() => {
    if (displayMode !== 'exact_docx') {
      lastRenderedSignatureRef.current = '';
    }
  }, [displayMode]);

  // Render exact exported .docx binary via `docx-preview` ONLY when base64Docx or replacements signature changes
  useEffect(() => {
    let cancelled = false;
    if (displayMode !== 'exact_docx' || !base64Docx || !docxHostRef.current) {
      return;
    }

    const renderSignature = `${base64Docx.length}:${base64Docx.slice(0, 64)}:${effectiveReplacementsSignature}`;
    if (
      lastRenderedSignatureRef.current === renderSignature &&
      docxHostRef.current.childElementCount > 0
    ) {
      return;
    }

    const hasExistingContent = docxHostRef.current.childElementCount > 0;
    const debounceMs = hasExistingContent ? 220 : 10;

    const timer = setTimeout(async () => {
      if (cancelled || !docxHostRef.current) return;
      setIsRenderingDocx(true);
      try {
        const currentReplacements = effectiveReplacementsRef.current;
        const currentParagraphs = paragraphsRef.current;
        const blob = await exportFilledDocxBlob(base64Docx, currentReplacements);
        if (cancelled || !docxHostRef.current) return;

        // Render into an offscreen staging container first to prevent layout collapse, white flash, or scroll jump
        const stagingContainer = document.createElement('div');
        await renderAsync(blob, stagingContainer, undefined, {
          className: 'docx',
          inWrapper: true,
          ignoreWidth: false,
          ignoreHeight: false,
          ignoreFonts: false,
          breakPages: true,
          renderHeaders: true,
          renderFooters: true,
          renderFootnotes: true,
          renderEndnotes: true,
          useBase64URL: true,
        });

        if (cancelled || !docxHostRef.current) return;

        // Tag rendered <p> nodes inside stagingContainer with data-docx-p-idx
        const renderedPs = Array.from(stagingContainer.querySelectorAll('p'));
        const usedParagraphIndices = new Set<number>();

        renderedPs.forEach((domP) => {
          const domText = (domP.textContent || '').trim();
          if (!domText) return;

          let matchedItem = currentParagraphs.find((item) => {
            if (usedParagraphIndices.has(item.index)) return false;
            const expectedText = (currentReplacements[item.index] ?? item.text).trim();
            return (
              expectedText === domText ||
              (expectedText.length > 4 && domText.includes(expectedText))
            );
          });

          if (!matchedItem) {
            matchedItem = currentParagraphs.find((item) => {
              const expectedText = (currentReplacements[item.index] ?? item.text).trim();
              return expectedText.length > 3 && domText.includes(expectedText);
            });
          }

          if (matchedItem) {
            const pIdx = matchedItem.index;
            usedParagraphIndices.add(pIdx);
            domP.setAttribute('data-docx-p-idx', String(pIdx));
            domP.setAttribute(
              'title',
              `[Dòng ${String(pIdx).padStart(2, '0')}] ${matchedItem.sectionLabel} — Nhấp để chỉnh sửa`
            );
            if (activeIndexRef.current === pIdx) {
              domP.setAttribute('data-docx-p-active', 'true');
            }
            if (
              currentReplacements[pIdx] !== undefined &&
              currentReplacements[pIdx] !== matchedItem.text
            ) {
              domP.setAttribute('data-docx-p-modified', 'true');
            }
          }
        });

        // Atomically replace DOM children while preserving scroll position
        const prevScrollTop = docxHostRef.current.scrollTop;
        const prevWindowScrollY = window.scrollY;
        docxHostRef.current.replaceChildren(...Array.from(stagingContainer.childNodes));
        docxHostRef.current.scrollTop = prevScrollTop;
        if (window.scrollY !== prevWindowScrollY) {
          window.scrollTo({ top: prevWindowScrollY });
        }
        lastRenderedSignatureRef.current = renderSignature;
      } catch (err) {
        console.error('docx-preview render error:', err);
      } finally {
        if (!cancelled) setIsRenderingDocx(false);
      }
    }, debounceMs);

    return () => {
      cancelled = true;
      clearTimeout(timer);
    };
  }, [displayMode, base64Docx, effectiveReplacementsSignature]);

  // Lightweight effect: update active line highlight in-place without re-rendering the .docx binary
  useEffect(() => {
    if (displayMode !== 'exact_docx' || !docxHostRef.current) return;
    const prevActive = docxHostRef.current.querySelectorAll('[data-docx-p-active]');
    prevActive.forEach((el) => el.removeAttribute('data-docx-p-active'));
    if (activeIndex !== null && activeIndex !== undefined) {
      const nextActive = docxHostRef.current.querySelectorAll(
        `[data-docx-p-idx="${activeIndex}"]`
      );
      nextActive.forEach((el) => el.setAttribute('data-docx-p-active', 'true'));
    }
  }, [displayMode, activeIndex]);

  // Single delegated click handler for docxHostRef
  const handleDocxHostClick = (e: React.MouseEvent<HTMLDivElement>) => {
    if (!onSelectParagraph) return;
    const targetEl = (e.target as HTMLElement | null)?.closest?.('[data-docx-p-idx]');
    if (!targetEl) return;
    const idxAttr = targetEl.getAttribute('data-docx-p-idx');
    if (idxAttr !== null) {
      const parsedIdx = Number(idxAttr);
      if (!Number.isNaN(parsedIdx)) {
        onSelectParagraph(parsedIdx);
      }
    }
  };

  // Group paragraphs into realistic Word A4 page zones: Header (Headnote), Body/Tables, Footnote/Endnote, and Footer
  const groupedSections = useMemo(() => {
    const headers: DocxParagraphItem[] = [];
    const bodies: DocxParagraphItem[] = [];
    const notes: DocxParagraphItem[] = [];
    const footers: DocxParagraphItem[] = [];

    paragraphs.forEach((p) => {
      if (p.section === 'header') {
        headers.push(p);
      } else if (p.section === 'footnote' || p.section === 'endnote') {
        notes.push(p);
      } else if (p.section === 'footer') {
        footers.push(p);
      } else {
        bodies.push(p);
      }
    });

    const bodyBlocks = groupBodyParagraphsAndTables(bodies);

    return { headers, bodies, bodyBlocks, notes, footers };
  }, [paragraphs]);

  const showTokensInline = displayMode === 'tokens_interactive';

  // Helper to render a paragraph that contains field tokens with interactive inline spans
  const renderParagraphWithMailingsTokens = (
    p: DocxParagraphItem,
    templateStr: string
  ): React.ReactNode => {
    const universalSource = buildUniversalFieldTokenRegex(fieldSymbol, false).source;
    const splitRegex = new RegExp(`(${universalSource})`, 'g');
    const parts = templateStr.split(splitRegex).filter((pt) => pt !== undefined && pt !== '');

    // Because capturing groups inside splitRegex can emit the inner key right after the full token match,
    // let's use matchAll / lastIndex slicing so we only get [plainText, fullToken, plainText, ...]
    const cleanSegments: Array<{ isToken: boolean; raw: string; fieldKey?: string }> = [];
    const matchRegex = buildUniversalFieldTokenRegex(fieldSymbol, true);
    let lastIdx = 0;
    let m: RegExpExecArray | null;
    while ((m = matchRegex.exec(templateStr)) !== null) {
      if (m.index > lastIdx) {
        cleanSegments.push({
          isToken: false,
          raw: templateStr.slice(lastIdx, m.index),
        });
      }
      let extractedKey: string | undefined;
      for (let g = 1; g < m.length; g++) {
        if (m[g]) {
          extractedKey = m[g];
          break;
        }
      }
      cleanSegments.push({
        isToken: true,
        raw: m[0],
        fieldKey: extractedKey || matchSingleFieldToken(m[0], fieldSymbol) || m[0],
      });
      lastIdx = m.index + m[0].length;
    }
    if (lastIdx < templateStr.length) {
      cleanSegments.push({
        isToken: false,
        raw: templateStr.slice(lastIdx),
      });
    }

    const firstRunBold = Boolean(
      p.runs && p.runs.length >= 2 && p.runs[0].bold && p.runs.some((r) => !r.bold)
    );

    return cleanSegments.map((seg, i) => {
      if (!seg.isToken || !seg.fieldKey) {
        const part = seg.raw;
        if (i === 0 && firstRunBold && part.includes(':')) {
          const colonPos = part.indexOf(':');
          const beforeColon = part.slice(0, colonPos + 1);
          const afterColon = part.slice(colonPos + 1);
          return (
            <React.Fragment key={i}>
              <span className="font-bold">{beforeColon}</span>
              <span className="font-normal">{afterColon}</span>
            </React.Fragment>
          );
        }
        return <React.Fragment key={i}>{part}</React.Fragment>;
      }

      const fieldKey = seg.fieldKey;
      const fieldDef = getFieldDefByKey(fieldKey, customMailingsFields);
      const currentVal = fieldValues[fieldKey];
      const isFieldFocused = activeFieldKey === fieldKey;
      const formattedToken = formatFieldToken(fieldKey, fieldSymbol);

      const displayText =
        !showTokensInline && mode === 'fill'
          ? currentVal !== undefined
            ? currentVal
            : ''
          : formattedToken;

      return (
        <span
          key={i}
          onClick={(e) => {
            if (onSelectFieldKey) {
              e.stopPropagation();
              onSelectFieldKey(fieldKey, p.index);
            }
          }}
          title={`Trường dữ liệu: ${formattedToken} (${fieldDef.label}) — Nhấp để sửa hoặc đổi trường`}
          className={`inline-block rounded px-1.5 py-0.5 mx-0.5 transition-colors cursor-pointer ${
            isFieldFocused
              ? 'bg-blue-700 text-white font-semibold font-sans text-[13.5px]'
              : showTokensInline || mode === 'configure'
              ? 'bg-blue-50 text-blue-800 border border-blue-300 font-mono-tabular font-medium text-[13px] hover:bg-blue-100'
              : 'bg-amber-50/90 text-slate-900 border-b-2 border-blue-600/70 font-normal hover:bg-blue-50'
          }`}
        >
          {displayText || '\u00A0\u00A0\u00A0'}
        </span>
      );
    });
  };

  // Render an individual paragraph block inside any zone (Header, Body, Table Cell, Footnote, Footer)
  const renderParagraphBlock = (p: DocxParagraphItem, compactInTable = false) => {
    const idxStr = String(p.index);
    const isSelected = selectedSet.has(p.index);
    const isActive = activeIndex === p.index;
    const pTpl = paragraphTemplates[idxStr];
    const hasMailingsTemplate = Boolean(pTpl && hasAnyMergeFieldToken(pTpl, fieldSymbol));

    // Check if the user manually edited the full line text so it diverges from the resolved Mailings template
    const resolvedFromFields = hasMailingsTemplate
      ? resolveParagraphTemplate(pTpl, fieldValues, false, fieldSymbol)
      : p.text;

    const currentText = effectiveReplacements[p.index] ?? p.text;
    const wasManuallyOverriddenInLineEditor =
      mode === 'fill' &&
      replacements[p.index] !== undefined &&
      hasMailingsTemplate &&
      replacements[p.index] !== resolvedFromFields;

    const isModified =
      mode === 'fill' &&
      effectiveReplacements[p.index] !== undefined &&
      effectiveReplacements[p.index] !== p.text;

    if (p.isEmpty && !isSelected && !isModified && !compactInTable) {
      return <div key={p.index} className="h-3" aria-hidden="true" />;
    }

    const alignClass =
      p.alignment === 'center'
        ? 'text-center'
        : p.alignment === 'right'
        ? 'text-right'
        : p.alignment === 'both'
        ? 'text-justify'
        : 'text-left';

    const weightClass = p.isBold ? 'font-bold' : 'font-normal';
    const styleClass = `${p.isItalic ? 'italic' : ''} ${
      p.isUnderline ? 'underline underline-offset-2' : ''
    }`;

    const isNote = p.section === 'footnote' || p.section === 'endnote';
    const isHeaderOrFooter = p.section === 'header' || p.section === 'footer';
    const isInteractive = Boolean(onSelectParagraph);

    return (
      <div
        key={p.index}
        onClick={() => {
          if (onSelectParagraph) {
            onSelectParagraph(p.index);
          }
        }}
        role={isInteractive ? 'button' : undefined}
        tabIndex={isInteractive ? 0 : undefined}
        onKeyDown={(e) => {
          if (onSelectParagraph && (e.key === 'Enter' || e.key === ' ')) {
            e.preventDefault();
            onSelectParagraph(p.index);
          }
        }}
        style={{
          paddingLeft: p.indentLeftPx ? `${p.indentLeftPx + 8}px` : undefined,
          textIndent: p.firstLineIndentPx ? `${p.firstLineIndentPx}px` : undefined,
        }}
        className={`group relative rounded ${
          compactInTable ? 'px-1.5 py-1' : 'px-2.5 py-1.5'
        } transition-colors duration-150 ${alignClass} ${
          isInteractive ? 'cursor-pointer' : ''
        } ${
          isActive
            ? 'bg-blue-50/90 border border-blue-600'
            : isModified
            ? 'bg-emerald-50/60 border border-emerald-500/50 hover:border-emerald-600'
            : isSelected
            ? 'bg-amber-50/50 border border-amber-400/60 hover:border-amber-500'
            : 'border border-transparent hover:bg-slate-50 hover:border-slate-300'
        }`}
      >
        {/* Line Index, Section / Table Coordinate & Field Meta Header inside Preview */}
        {(showLineNumbers || isSelected || isModified || mode === 'configure') && (
          <div
            style={{ textIndent: 0 }}
            className={`flex flex-wrap items-center justify-between gap-2 mb-0.5 font-sans text-[11px] leading-tight select-none ${
              isActive
                ? 'text-blue-800'
                : isModified
                ? 'text-emerald-800'
                : isSelected
                ? 'text-amber-900'
                : 'text-slate-400 opacity-75 group-hover:opacity-100'
            }`}
          >
            <div className="flex flex-wrap items-center gap-1.5">
              <span className="font-mono-tabular font-medium">
                [Dòng {String(p.index).padStart(2, '0')}]
              </span>
              {p.section !== 'body' && (
                <>
                  <span aria-hidden="true">·</span>
                  <span className="font-semibold text-blue-700">
                    {p.sectionLabel}
                  </span>
                </>
              )}
              {p.tablePosition && (
                <>
                  <span aria-hidden="true">·</span>
                  <span className="font-mono-tabular font-medium text-slate-600">
                    {p.tablePosition}
                  </span>
                </>
              )}
              {customLabels[idxStr] && (
                <>
                  <span aria-hidden="true">·</span>
                  <span className="font-medium text-slate-700">
                    {customLabels[idxStr]}
                  </span>
                </>
              )}
            </div>
            <span className="font-sans text-[11px] font-medium">
              {isActive && mode === 'configure'
                ? 'Đang chọn vị trí chèn trường'
                : isActive && mode === 'fill'
                ? 'Đang sửa dòng này'
                : hasMailingsTemplate && !wasManuallyOverriddenInLineEditor
                ? 'Trường Mailings'
                : isModified
                ? 'Đã chỉnh sửa'
                : isSelected
                ? mode === 'fill'
                  ? 'Dòng nhập liệu'
                  : 'Đã đánh dấu'
                : 'Nhấp để sửa'}
            </span>
          </div>
        )}

        {/* Actual Document Paragraph Text or Inline Mailings Tokens */}
        <p
          style={{
            fontSize: p.fontSizePt ? `${p.fontSizePt}pt` : undefined,
          }}
          className={`${
            isNote
              ? 'text-[10.5pt] leading-snug text-slate-800'
              : isHeaderOrFooter
              ? 'text-[10pt] leading-snug text-slate-700'
              : 'text-[13pt] leading-relaxed text-slate-900'
          } whitespace-pre-wrap break-words ${weightClass} ${styleClass}`}
        >
          {/* If this paragraph is a Footnote or Endnote, display its superscript number */}
          {isNote && p.noteId && (
            <sup className="font-mono-tabular font-bold text-blue-700 mr-1.5 text-[10px]">
              ({p.noteId})
            </sup>
          )}

          {hasMailingsTemplate &&
          (showTokensInline || mode === 'configure' || !wasManuallyOverriddenInLineEditor)
            ? renderParagraphWithMailingsTokens(p, pTpl)
            : renderStyledTextFromOriginalRuns(p.text, currentText, p.runs)}

          {/* If this body paragraph references a Footnote in the Word XML, show clickable superscript reference */}
          {p.footnoteRefs &&
            p.footnoteRefs.length > 0 &&
            p.footnoteRefs.map((refId) => {
              const targetNote = groupedSections.notes.find((n) => n.noteId === refId);
              return (
                <sup
                  key={refId}
                  onClick={(e) => {
                    if (targetNote && onSelectParagraph) {
                      e.stopPropagation();
                      onSelectParagraph(targetNote.index);
                    }
                  }}
                  title={
                    targetNote
                      ? `Chú thích chân trang (Footnote #${refId}): ${targetNote.text} — Nhấp để chỉnh sửa Footnote`
                      : `Chú thích chân trang (Footnote #${refId})`
                  }
                  className="inline-block ml-1 px-1 py-0.2 text-[10px] font-sans font-bold text-blue-700 bg-blue-50 border border-blue-200 rounded hover:bg-blue-100 cursor-pointer"
                >
                  [{refId}]
                </sup>
              );
            })}
        </p>
      </div>
    );
  };

  // Render a real Word Table block (<w:tbl>) with rows and columns
  const renderTableBlock = (tableBlock: TableGroupBlock) => {
    const rowEntries = Array.from(tableBlock.rows.entries()).sort((a, b) => a[0] - b[0]);
    return (
      <div key={`table-${tableBlock.tableId}`} className="my-2 overflow-x-auto">
        <table
          className={`w-full border-collapse ${
            tableBlock.hasBorders ? 'border border-slate-400' : 'border border-dashed border-slate-300'
          }`}
        >
          <tbody>
            {rowEntries.map(([rIdx, colMap]) => {
              const colEntries = Array.from(colMap.entries()).sort((a, b) => a[0] - b[0]);
              return (
                <tr key={rIdx} className="align-top">
                  {colEntries.map(([cIdx, cell]) => (
                    <td
                      key={cIdx}
                      colSpan={cell.colSpan > 1 ? cell.colSpan : undefined}
                      className={`p-1.5 align-top ${
                        tableBlock.hasBorders
                          ? 'border border-slate-400'
                          : 'border border-dashed border-slate-200'
                      }`}
                    >
                      <div className="space-y-1">
                        {cell.paragraphs.map((p) => renderParagraphBlock(p, true))}
                      </div>
                    </td>
                  ))}
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
    );
  };

  return (
    <div className="flex flex-col h-full">
      {/* Word Mailings Preview Toolbar */}
      <div className="flex flex-wrap items-center justify-between gap-3 px-4 py-3 bg-slate-100 border border-slate-200 border-b-0 rounded-t-lg">
        <div className="flex flex-wrap items-center gap-2 text-xs text-slate-600 min-w-0">
          <span className="font-semibold text-slate-900 truncate">
            {templateName || 'Trang văn bản Word A4'}
          </span>
          <span aria-hidden="true">·</span>
          <span className="font-mono-tabular whitespace-nowrap">
            {paragraphs.filter((p) => !p.isEmpty).length} dòng
          </span>
          {groupedSections.headers.length > 0 && (
            <>
              <span aria-hidden="true">·</span>
              <span className="font-mono-tabular text-blue-800 whitespace-nowrap">
                {groupedSections.headers.length} Headnote
              </span>
            </>
          )}
          {groupedSections.notes.length > 0 && (
            <>
              <span aria-hidden="true">·</span>
              <span className="font-mono-tabular text-blue-800 whitespace-nowrap">
                {groupedSections.notes.length} Footnote
              </span>
            </>
          )}
        </div>

        {/* 3-Mode Preview Switcher: Mã «Trường» | Tương tác dòng | Chuẩn File Word Xuất Ra (.docx) */}
        <div className="flex flex-wrap items-center gap-1 p-0.5 bg-slate-200/80 rounded-md shrink-0">
          <button
            type="button"
            onClick={() => setDisplayMode('tokens_interactive')}
            className={`inline-flex items-center gap-1 px-2.5 py-1 text-[11px] font-medium rounded transition-colors whitespace-nowrap shrink-0 cursor-pointer ${
              displayMode === 'tokens_interactive'
                ? 'bg-white text-slate-900 shadow-2xs font-semibold'
                : 'text-slate-600 hover:text-slate-900'
            }`}
          >
            <Code2 className="w-3 h-3 text-blue-700" />
            <span>Hiện mã «Trường»</span>
          </button>
          <button
            type="button"
            onClick={() => setDisplayMode('merged_interactive')}
            className={`inline-flex items-center gap-1 px-2.5 py-1 text-[11px] font-medium rounded transition-colors whitespace-nowrap shrink-0 cursor-pointer ${
              displayMode === 'merged_interactive'
                ? 'bg-white text-slate-900 shadow-2xs font-semibold'
                : 'text-slate-600 hover:text-slate-900'
            }`}
          >
            <Eye className="w-3 h-3 text-blue-700" />
            <span>Tương tác theo dòng</span>
          </button>
          {base64Docx && (
            <button
              type="button"
              onClick={() => setDisplayMode('exact_docx')}
              className={`inline-flex items-center gap-1 px-2.5 py-1 text-[11px] font-medium rounded transition-colors whitespace-nowrap shrink-0 cursor-pointer ${
                displayMode === 'exact_docx'
                  ? 'bg-blue-700 text-white shadow-2xs font-semibold'
                  : 'text-slate-700 hover:text-slate-900'
              }`}
            >
              <FileCheck2 className="w-3 h-3" />
              <span>Bản Word xuất ra (100% khớp .docx)</span>
            </button>
          )}
        </div>
      </div>

      {/* Mode 1: Exact .docx Binary Renderer via docx-preview (100% identical to exported Word file) */}
      {displayMode === 'exact_docx' && base64Docx ? (
        <div className="a4-paper-preview w-[700px] rounded-b-lg overflow-hidden">
          <div className="px-4 py-2 bg-blue-50/80 border-b border-blue-200 flex items-center justify-between gap-2 text-xs text-blue-900">
            <span>
              Đang hiển thị <strong>trực tiếp từ file .docx xuất ra</strong> (giữ nguyên 100% bảng biểu, căn lề, phông chữ, in đậm/nghiêng, Headnote, Footnote, Footer). Nhấp vào bất kỳ dòng nào trên trang Word để chỉnh sửa.
            </span>
            {isRenderingDocx && (
              <span className="font-mono-tabular text-[11px] text-blue-700 shrink-0">
                Đang kết xuất .docx...
              </span>
            )}
          </div>
          <div
            ref={docxHostRef}
            onClick={handleDocxHostClick}
            className="docx-preview-host min-h-[640px]"
          />
        </div>
      ) : (
        /* Mode 2 & 3: Interactive Line & Mailings Sheet Canvas (with real Tables & Run formatting) */
        <div className="a4-paper-preview w-[700px] rounded-b-lg p-6 sm:p-10 lg:p-12 min-h-[640px] flex flex-col justify-between transition-opacity duration-150">
          {paragraphs.length === 0 ? (
            <div className="py-24 text-center text-sm text-slate-400">
              Tài liệu chưa có nội dung văn bản để hiển thị. Hãy tải lên file Word (.docx) ở khung bên trái.
            </div>
          ) : (
            <div
              className="flex flex-col flex-1 justify-between space-y-6 text-slate-900"
              style={{ fontFamily: '"Times New Roman", Times, serif' }}
            >
              {/* Top + Middle Area: Header (Headnote) & Body / Tables */}
              <div className="space-y-5">
                {/* 1. WORD HEADER (HEADNOTE / ĐẦU TRANG) ZONE */}
                {groupedSections.headers.length > 0 && (
                  <div className="pb-3.5 border-b border-dashed border-slate-300 space-y-1.5">
                    <div className="flex items-center justify-between font-sans text-[10px] uppercase tracking-wider font-semibold text-slate-400 px-1 select-none">
                      <span>Đầu trang (Headnote / Header Word)</span>
                      <span>Nhấp vào dòng để chỉnh sửa Headnote</span>
                    </div>
                    {groupedSections.headers.map((p) => renderParagraphBlock(p))}
                  </div>
                )}

                {/* 2. WORD BODY & TABLES ZONE (renders real tables & paragraphs) */}
                <div className="space-y-2.5">
                  {groupedSections.bodyBlocks.map((block) =>
                    block.kind === 'table'
                      ? renderTableBlock(block)
                      : renderParagraphBlock(block.paragraph)
                  )}
                </div>
              </div>

              {/* Bottom Area: Footnotes / Endnotes & Footer */}
              {(groupedSections.notes.length > 0 || groupedSections.footers.length > 0) && (
                <div className="pt-6 space-y-4">
                  {/* 3. WORD FOOTNOTES & ENDNOTES (CHÚ THÍCH CHÂN TRANG) ZONE */}
                  {groupedSections.notes.length > 0 && (
                    <div className="space-y-2">
                      {/* Classic Word Footnote short horizontal rule */}
                      <div className="w-44 border-t border-slate-400 mb-2" aria-hidden="true" />
                      <div className="flex items-center justify-between font-sans text-[10px] uppercase tracking-wider font-semibold text-blue-800/80 px-1 select-none">
                        <span>Chú thích chân trang (Footnote / Endnote Word)</span>
                        <span>Nhấp vào dòng để chỉnh sửa Footnote</span>
                      </div>
                      <div className="space-y-1.5">
                        {groupedSections.notes.map((p) => renderParagraphBlock(p))}
                      </div>
                    </div>
                  )}

                  {/* 4. WORD FOOTER (CHÂN TRANG) ZONE */}
                  {groupedSections.footers.length > 0 && (
                    <div className="pt-3 border-t border-dashed border-slate-300 space-y-1.5">
                      <div className="flex items-center justify-between font-sans text-[10px] uppercase tracking-wider font-semibold text-slate-400 px-1 select-none">
                        <span>Chân trang (Footer Word)</span>
                        <span>Nhấp vào dòng để chỉnh sửa Footer</span>
                      </div>
                      {groupedSections.footers.map((p) => renderParagraphBlock(p))}
                    </div>
                  )}
                </div>
              )}
            </div>
          )}
        </div>
      )}
    </div>
  );
};
