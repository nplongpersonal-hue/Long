import JSZip from 'jszip';

export type DocxSectionType = 'header' | 'body' | 'footnote' | 'endnote' | 'footer';

export interface DocxRunStyle {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  fontSizePt?: number;
  color?: string;
}

export interface DocxRunSpan extends DocxRunStyle {
  text: string;
}

export interface DocxParagraphItem {
  index: number;
  text: string;
  isEmpty: boolean;
  section: DocxSectionType;
  sectionLabel: string;
  xmlPath: string;
  noteId?: string;
  footnoteRefs?: string[];
  alignment?: 'left' | 'center' | 'right' | 'both';
  isBold?: boolean;
  isItalic?: boolean;
  isUnderline?: boolean;
  fontSizePt?: number;
  indentLeftPx?: number;
  firstLineIndentPx?: number;
  runs?: DocxRunSpan[];
  isInTable?: boolean;
  tableId?: number;
  tableRowIndex?: number;
  tableColIndex?: number;
  tableTotalCols?: number;
  tableColSpan?: number;
  tableHasBorders?: boolean;
  tablePosition?: string;
}

export interface ParsedDocxResult {
  allParagraphs: DocxParagraphItem[];
  nonEmptyParagraphs: DocxParagraphItem[];
}

const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

export function base64ToUint8Array(base64: string): Uint8Array {
  const binaryString = window.atob(base64);
  const len = binaryString.length;
  const bytes = new Uint8Array(len);
  for (let i = 0; i < len; i++) {
    bytes[i] = binaryString.charCodeAt(i);
  }
  return bytes;
}

export function arrayBufferToBase64(buffer: ArrayBuffer): string {
  const bytes = new Uint8Array(buffer);
  let binary = '';
  const chunkSize = 0x8000;
  for (let i = 0; i < bytes.length; i += chunkSize) {
    const chunk = bytes.subarray(i, i + chunkSize);
    binary += String.fromCharCode.apply(null, Array.from(chunk));
  }
  return window.btoa(binary);
}

interface XmlParagraphTarget {
  xmlPath: string;
  xmlDoc: Document;
  pNode: Element;
  section: DocxSectionType;
  sectionLabel: string;
  noteId?: string;
}

/**
 * Collects all <w:p> elements inside a root container without duplicating paragraphs inside fallback shapes.
 */
function collectParagraphElementsInOrder(container: Element): Element[] {
  const results: Element[] = [];
  const walker = document.createTreeWalker(container, NodeFilter.SHOW_ELEMENT);
  let current = walker.nextNode() as Element | null;
  while (current) {
    if (current.localName === 'p' || current.nodeName === 'w:p') {
      let insideFallback = false;
      let parent = current.parentElement;
      while (parent && parent !== container) {
        if (parent.localName === 'Fallback' || parent.nodeName === 'mc:Fallback') {
          insideFallback = true;
          break;
        }
        parent = parent.parentElement;
      }
      if (!insideFallback) {
        results.push(current);
      }
    }
    current = walker.nextNode() as Element | null;
  }
  return results;
}

/**
 * Collects all editable/viewable paragraph nodes across the entire .docx package in a deterministic order:
 * 1. Headers (word/header*.xml) — Headnotes / Tiêu đề đầu trang
 * 2. Document Body & Tables (word/document.xml) — Nội dung chính & Bảng biểu
 * 3. Footnotes (word/footnotes.xml) — Chú thích chân trang
 * 4. Endnotes (word/endnotes.xml) — Chú thích cuối tài liệu
 * 5. Footers (word/footer*.xml) — Chân trang
 */
async function collectAllDocxParagraphTargets(zip: JSZip): Promise<{
  targets: XmlParagraphTarget[];
  xmlDocsByPath: Map<string, Document>;
}> {
  const parser = new DOMParser();
  const targets: XmlParagraphTarget[] = [];
  const xmlDocsByPath = new Map<string, Document>();

  const allFileNames = Object.keys(zip.files).sort((a, b) => a.localeCompare(b));

  // 1. Headers (word/header1.xml, word/header2.xml, ...)
  const headerPaths = allFileNames.filter((name) =>
    /^word\/header\d*\.xml$/i.test(name)
  );
  for (let hIdx = 0; hIdx < headerPaths.length; hIdx++) {
    const xmlPath = headerPaths[hIdx];
    const fileObj = zip.file(xmlPath);
    if (!fileObj) continue;
    const xmlText = await fileObj.async('string');
    const xmlDoc = parser.parseFromString(xmlText, 'application/xml');
    xmlDocsByPath.set(xmlPath, xmlDoc);

    const pNodes = collectParagraphElementsInOrder(xmlDoc.documentElement);
    const label =
      headerPaths.length > 1
        ? `Đầu trang (Headnote / Header #${hIdx + 1})`
        : 'Đầu trang (Headnote / Header)';

    pNodes.forEach((pNode) => {
      targets.push({
        xmlPath,
        xmlDoc,
        pNode,
        section: 'header',
        sectionLabel: label,
      });
    });
  }

  // 2. Main Document Body (word/document.xml)
  const docXmlFile = zip.file('word/document.xml');
  if (!docXmlFile) {
    throw new Error('File không đúng chuẩn Microsoft Word (.docx): Thiếu word/document.xml');
  }
  const docXmlText = await docXmlFile.async('string');
  const docXml = parser.parseFromString(docXmlText, 'application/xml');
  xmlDocsByPath.set('word/document.xml', docXml);

  const bodyNode =
    docXml.getElementsByTagNameNS(W_NS, 'body')[0] ||
    docXml.getElementsByTagName('w:body')[0];

  if (!bodyNode) {
    throw new Error('Không tìm thấy phần thân tài liệu (w:body) trong file Word.');
  }

  const bodyPNodes = collectParagraphElementsInOrder(bodyNode);
  bodyPNodes.forEach((pNode) => {
    targets.push({
      xmlPath: 'word/document.xml',
      xmlDoc: docXml,
      pNode,
      section: 'body',
      sectionLabel: 'Nội dung chính (Body)',
    });
  });

  // 3. Footnotes (word/footnotes.xml)
  const footnotesFile = zip.file('word/footnotes.xml');
  if (footnotesFile) {
    const fnXmlText = await footnotesFile.async('string');
    const fnXmlDoc = parser.parseFromString(fnXmlText, 'application/xml');
    xmlDocsByPath.set('word/footnotes.xml', fnXmlDoc);

    const footnoteElems = Array.from(fnXmlDoc.documentElement.childNodes).filter(
      (n): n is Element =>
        n.nodeType === 1 &&
        ((n as Element).localName === 'footnote' ||
          (n as Element).nodeName === 'w:footnote')
    );

    footnoteElems.forEach((fnElem) => {
      const fnType =
        fnElem.getAttribute('w:type') || fnElem.getAttributeNS(W_NS, 'type') || '';
      const fnId =
        fnElem.getAttribute('w:id') || fnElem.getAttributeNS(W_NS, 'id') || '';

      if (
        fnType === 'separator' ||
        fnType === 'continuationSeparator' ||
        fnId === '0' ||
        fnId === '-1'
      ) {
        return;
      }

      const pNodes = collectParagraphElementsInOrder(fnElem);
      pNodes.forEach((pNode) => {
        targets.push({
          xmlPath: 'word/footnotes.xml',
          xmlDoc: fnXmlDoc,
          pNode,
          section: 'footnote',
          noteId: fnId,
          sectionLabel: `Chú thích chân trang (Footnote #${fnId})`,
        });
      });
    });
  }

  // 4. Endnotes (word/endnotes.xml)
  const endnotesFile = zip.file('word/endnotes.xml');
  if (endnotesFile) {
    const enXmlText = await endnotesFile.async('string');
    const enXmlDoc = parser.parseFromString(enXmlText, 'application/xml');
    xmlDocsByPath.set('word/endnotes.xml', enXmlDoc);

    const endnoteElems = Array.from(enXmlDoc.documentElement.childNodes).filter(
      (n): n is Element =>
        n.nodeType === 1 &&
        ((n as Element).localName === 'endnote' ||
          (n as Element).nodeName === 'w:endnote')
    );

    endnoteElems.forEach((enElem) => {
      const enType =
        enElem.getAttribute('w:type') || enElem.getAttributeNS(W_NS, 'type') || '';
      const enId =
        enElem.getAttribute('w:id') || enElem.getAttributeNS(W_NS, 'id') || '';

      if (
        enType === 'separator' ||
        enType === 'continuationSeparator' ||
        enId === '0' ||
        enId === '-1'
      ) {
        return;
      }

      const pNodes = collectParagraphElementsInOrder(enElem);
      pNodes.forEach((pNode) => {
        targets.push({
          xmlPath: 'word/endnotes.xml',
          xmlDoc: enXmlDoc,
          pNode,
          section: 'endnote',
          noteId: enId,
          sectionLabel: `Chú thích cuối (Endnote #${enId})`,
        });
      });
    });
  }

  // 5. Footers (word/footer1.xml, word/footer2.xml, ...)
  const footerPaths = allFileNames.filter((name) =>
    /^word\/footer\d*\.xml$/i.test(name)
  );
  for (let fIdx = 0; fIdx < footerPaths.length; fIdx++) {
    const xmlPath = footerPaths[fIdx];
    const fileObj = zip.file(xmlPath);
    if (!fileObj) continue;
    const xmlText = await fileObj.async('string');
    const xmlDoc = parser.parseFromString(xmlText, 'application/xml');
    xmlDocsByPath.set(xmlPath, xmlDoc);

    const pNodes = collectParagraphElementsInOrder(xmlDoc.documentElement);
    const label =
      footerPaths.length > 1
        ? `Chân trang (Footer #${fIdx + 1})`
        : 'Chân trang (Footer)';

    pNodes.forEach((pNode) => {
      targets.push({
        xmlPath,
        xmlDoc,
        pNode,
        section: 'footer',
        sectionLabel: label,
      });
    });
  }

  return { targets, xmlDocsByPath };
}

export async function parseDocxFromBase64(base64Data: string): Promise<ParsedDocxResult> {
  const bytes = base64ToUint8Array(base64Data);
  const zip = await JSZip.loadAsync(bytes);

  const { targets } = await collectAllDocxParagraphTargets(zip);

  // Track unique <w:tbl> elements to assign deterministic tableId
  const seenTables: Element[] = [];

  const allParagraphs: DocxParagraphItem[] = targets.map((target, idx) => {
    const { text, footnoteRefs, runs } = extractParagraphTextRunsAndRefs(target.pNode);
    const alignment = detectAlignment(target.pNode);
    const { indentLeftPx, firstLineIndentPx } = detectIndentation(target.pNode);
    const { isBold, isItalic, isUnderline, fontSizePt } = detectParagraphSummaryStyle(target.pNode, runs);
    const tableMeta = detectTableMetadata(target.pNode, seenTables);

    return {
      index: idx,
      text,
      isEmpty: text.trim().length === 0,
      section: target.section,
      sectionLabel: tableMeta.isInTable
        ? `${target.sectionLabel} · ${tableMeta.tablePosition}`
        : target.sectionLabel,
      xmlPath: target.xmlPath,
      noteId: target.noteId,
      footnoteRefs: footnoteRefs.length > 0 ? footnoteRefs : undefined,
      alignment,
      isBold,
      isItalic,
      isUnderline,
      fontSizePt,
      indentLeftPx,
      firstLineIndentPx,
      runs: runs.length > 0 ? runs : undefined,
      isInTable: tableMeta.isInTable,
      tableId: tableMeta.tableId,
      tableRowIndex: tableMeta.tableRowIndex,
      tableColIndex: tableMeta.tableColIndex,
      tableTotalCols: tableMeta.tableTotalCols,
      tableColSpan: tableMeta.tableColSpan,
      tableHasBorders: tableMeta.tableHasBorders,
      tablePosition: tableMeta.tablePosition,
    };
  });

  const nonEmptyParagraphs = allParagraphs.filter((p) => !p.isEmpty);

  return {
    allParagraphs,
    nonEmptyParagraphs,
  };
}

function parseRPrStyle(rPr: Element | null): DocxRunStyle {
  if (!rPr) return {};
  const b = findFirstChildByLocalName(rPr, 'b');
  const i = findFirstChildByLocalName(rPr, 'i');
  const u = findFirstChildByLocalName(rPr, 'u');
  const sz = findFirstChildByLocalName(rPr, 'sz');
  const color = findFirstChildByLocalName(rPr, 'color');

  const isBold = Boolean(
    b && b.getAttribute('w:val') !== '0' && b.getAttribute('w:val') !== 'false'
  );
  const isItalic = Boolean(
    i && i.getAttribute('w:val') !== '0' && i.getAttribute('w:val') !== 'false'
  );
  const uVal = u ? u.getAttribute('w:val') || u.getAttributeNS(W_NS, 'val') : null;
  const isUnderline = Boolean(u && uVal !== 'none' && uVal !== '0' && uVal !== 'false');

  let fontSizePt: number | undefined;
  if (sz) {
    const rawSz = sz.getAttribute('w:val') || sz.getAttributeNS(W_NS, 'val');
    if (rawSz && !Number.isNaN(Number(rawSz))) {
      fontSizePt = Number(rawSz) / 2;
    }
  }

  let hexColor: string | undefined;
  if (color) {
    const rawColor = color.getAttribute('w:val') || color.getAttributeNS(W_NS, 'val');
    if (rawColor && rawColor !== 'auto' && /^[0-9A-Fa-f]{6}$/.test(rawColor)) {
      hexColor = `#${rawColor}`;
    }
  }

  return {
    bold: isBold || undefined,
    italic: isItalic || undefined,
    underline: isUnderline || undefined,
    fontSizePt,
    color: hexColor,
  };
}

function findAncestorRun(node: Element, stopAt: Element): Element | null {
  let cur: Element | null = node;
  while (cur && cur !== stopAt) {
    if (cur.localName === 'r' || cur.nodeName === 'w:r') {
      return cur;
    }
    cur = cur.parentElement;
  }
  return null;
}

function extractParagraphTextRunsAndRefs(pNode: Element): {
  text: string;
  footnoteRefs: string[];
  runs: DocxRunSpan[];
  charRPrs: (Element | null)[];
} {
  const texts: string[] = [];
  const footnoteRefs: string[] = [];
  const runs: DocxRunSpan[] = [];
  const charRPrs: (Element | null)[] = [];

  const pPr = findFirstChildByLocalName(pNode, 'pPr');
  const pDefaultRPr = pPr ? findFirstChildByLocalName(pPr, 'rPr') : null;

  const appendPiece = (str: string, rNode: Element | null) => {
    if (!str) return;
    texts.push(str);
    const rPr =
      (rNode ? findFirstChildByLocalName(rNode, 'rPr') : null) || pDefaultRPr || null;
    for (let k = 0; k < str.length; k++) {
      charRPrs.push(rPr);
    }
    const style = parseRPrStyle(rPr);
    const prev = runs[runs.length - 1];
    if (
      prev &&
      prev.bold === style.bold &&
      prev.italic === style.italic &&
      prev.underline === style.underline &&
      prev.fontSizePt === style.fontSizePt &&
      prev.color === style.color
    ) {
      prev.text += str;
    } else {
      runs.push({ text: str, ...style });
    }
  };

  const walker = document.createTreeWalker(pNode, NodeFilter.SHOW_ELEMENT);
  let current = walker.nextNode() as Element | null;
  while (current) {
    const local = current.localName;
    const nodeName = current.nodeName;
    if (local === 't' || nodeName === 'w:t' || local === 'delText' || nodeName === 'w:delText') {
      const rNode = findAncestorRun(current, pNode);
      appendPiece(current.textContent || '', rNode);
    } else if (local === 'tab' || nodeName === 'w:tab') {
      const rNode = findAncestorRun(current, pNode);
      appendPiece('\t', rNode);
    } else if (local === 'br' || nodeName === 'w:br') {
      const rNode = findAncestorRun(current, pNode);
      appendPiece('\n', rNode);
    } else if (local === 'footnoteReference' || nodeName === 'w:footnoteReference') {
      const fnId =
        current.getAttribute('w:id') || current.getAttributeNS(W_NS, 'id') || '';
      if (fnId && fnId !== '0' && fnId !== '-1') {
        footnoteRefs.push(fnId);
      }
    } else if (local === 'endnoteReference' || nodeName === 'w:endnoteReference') {
      const enId =
        current.getAttribute('w:id') || current.getAttributeNS(W_NS, 'id') || '';
      if (enId && enId !== '0' && enId !== '-1') {
        footnoteRefs.push(`E${enId}`);
      }
    }
    current = walker.nextNode() as Element | null;
  }

  return {
    text: texts.join(''),
    footnoteRefs,
    runs,
    charRPrs,
  };
}

function detectAlignment(pNode: Element): 'left' | 'center' | 'right' | 'both' {
  const pPr = findFirstChildByLocalName(pNode, 'pPr');
  if (!pPr) return 'left';
  const jc = findFirstChildByLocalName(pPr, 'jc');
  if (!jc) return 'left';
  const val = jc.getAttribute('w:val') || jc.getAttributeNS(W_NS, 'val') || 'left';
  if (val === 'center') return 'center';
  if (val === 'right' || val === 'end') return 'right';
  if (val === 'both' || val === 'distribute') return 'both';
  return 'left';
}

function detectIndentation(pNode: Element): {
  indentLeftPx?: number;
  firstLineIndentPx?: number;
} {
  const pPr = findFirstChildByLocalName(pNode, 'pPr');
  if (!pPr) return {};
  const ind = findFirstChildByLocalName(pPr, 'ind');
  if (!ind) return {};
  const leftTwips = Number(
    ind.getAttribute('w:left') || ind.getAttributeNS(W_NS, 'left') || '0'
  );
  const firstLineTwips = Number(
    ind.getAttribute('w:firstLine') || ind.getAttributeNS(W_NS, 'firstLine') || '0'
  );
  // 1 twip = 1/20 pt ≈ 1/15 px
  return {
    indentLeftPx: leftTwips > 0 ? Math.round(leftTwips / 15) : undefined,
    firstLineIndentPx: firstLineTwips > 0 ? Math.round(firstLineTwips / 15) : undefined,
  };
}

function detectParagraphSummaryStyle(
  pNode: Element,
  runs: DocxRunSpan[]
): {
  isBold: boolean;
  isItalic: boolean;
  isUnderline: boolean;
  fontSizePt?: number;
} {
  const nonEmptyRuns = runs.filter((r) => r.text.trim().length > 0);
  if (nonEmptyRuns.length > 0) {
    // Only mark the whole paragraph as bold/italic/underline if ALL non-empty runs have it
    const allBold = nonEmptyRuns.every((r) => Boolean(r.bold));
    const allItalic = nonEmptyRuns.every((r) => Boolean(r.italic));
    const allUnderline = nonEmptyRuns.every((r) => Boolean(r.underline));
    const fontSizePt = nonEmptyRuns[0].fontSizePt;
    return {
      isBold: allBold,
      isItalic: allItalic,
      isUnderline: allUnderline,
      fontSizePt,
    };
  }
  const firstRun = findFirstChildByLocalName(pNode, 'r');
  const rPr = firstRun ? findFirstChildByLocalName(firstRun, 'rPr') : null;
  const st = parseRPrStyle(rPr);
  return {
    isBold: Boolean(st.bold),
    isItalic: Boolean(st.italic),
    isUnderline: Boolean(st.underline),
    fontSizePt: st.fontSizePt,
  };
}

function detectTableMetadata(
  el: Element,
  seenTables: Element[]
): {
  isInTable: boolean;
  tableId?: number;
  tableRowIndex?: number;
  tableColIndex?: number;
  tableTotalCols?: number;
  tableColSpan?: number;
  tableHasBorders?: boolean;
  tablePosition?: string;
} {
  let parent = el.parentElement;
  let tc: Element | null = null;
  let tr: Element | null = null;
  let tbl: Element | null = null;

  while (parent) {
    const local = parent.localName || parent.nodeName;
    if (!tc && (local === 'tc' || local === 'w:tc')) tc = parent;
    if (!tr && (local === 'tr' || local === 'w:tr')) tr = parent;
    if (!tbl && (local === 'tbl' || local === 'w:tbl')) {
      tbl = parent;
      break;
    }
    parent = parent.parentElement;
  }

  if (!tbl) {
    return { isInTable: false };
  }

  let tableId = seenTables.indexOf(tbl);
  if (tableId === -1) {
    seenTables.push(tbl);
    tableId = seenTables.length - 1;
  }

  let rowIdx = 1;
  let colIdx = 1;
  let totalCols = 1;
  let colSpan = 1;

  if (tr && tbl) {
    const rows = Array.from(tbl.childNodes).filter(
      (n) => n.nodeType === 1 && ((n as Element).localName === 'tr' || (n as Element).nodeName === 'w:tr')
    );
    const foundRow = rows.indexOf(tr);
    if (foundRow !== -1) rowIdx = foundRow + 1;
  }
  if (tc && tr) {
    const cells = Array.from(tr.childNodes).filter(
      (n) => n.nodeType === 1 && ((n as Element).localName === 'tc' || (n as Element).nodeName === 'w:tc')
    );
    totalCols = Math.max(1, cells.length);
    const foundCol = cells.indexOf(tc);
    if (foundCol !== -1) colIdx = foundCol + 1;

    const tcPr = findFirstChildByLocalName(tc, 'tcPr');
    if (tcPr) {
      const gridSpan = findFirstChildByLocalName(tcPr, 'gridSpan');
      if (gridSpan) {
        const spanVal = Number(
          gridSpan.getAttribute('w:val') || gridSpan.getAttributeNS(W_NS, 'val') || '1'
        );
        if (!Number.isNaN(spanVal) && spanVal > 1) {
          colSpan = spanVal;
        }
      }
    }
  }

  // Check if table or cell has visible borders
  let tableHasBorders = true;
  const tblPr = findFirstChildByLocalName(tbl, 'tblPr');
  if (tblPr) {
    const tblBorders = findFirstChildByLocalName(tblPr, 'tblBorders');
    if (tblBorders) {
      const borderNodes = Array.from(tblBorders.childNodes).filter((n) => n.nodeType === 1) as Element[];
      if (borderNodes.length > 0) {
        const allNil = borderNodes.every((b) => {
          const val = b.getAttribute('w:val') || b.getAttributeNS(W_NS, 'val');
          return val === 'none' || val === 'nil';
        });
        if (allNil) tableHasBorders = false;
      }
    }
  }

  return {
    isInTable: true,
    tableId,
    tableRowIndex: rowIdx,
    tableColIndex: colIdx,
    tableTotalCols: totalCols,
    tableColSpan: colSpan,
    tableHasBorders,
    tablePosition: `Bảng (Hàng ${rowIdx}, Cột ${colIdx})`,
  };
}

function findFirstChildByLocalName(parent: Element, localName: string): Element | null {
  for (let i = 0; i < parent.childNodes.length; i++) {
    const child = parent.childNodes[i] as Element;
    if (child.nodeType === 1 && (child.localName === localName || child.nodeName === `w:${localName}`)) {
      return child;
    }
  }
  return null;
}

/**
 * Checks if a <w:r> element contains a footnoteRef, endnoteRef, footnoteReference, or endnoteReference
 * so that replacing paragraph text preserves footnote numbering markers in Word.
 */
function runHasNoteMarker(rNode: Element): boolean {
  const walker = document.createTreeWalker(rNode, NodeFilter.SHOW_ELEMENT);
  let cur = walker.nextNode() as Element | null;
  while (cur) {
    const local = cur.localName || cur.nodeName;
    if (
      local === 'footnoteRef' ||
      local === 'w:footnoteRef' ||
      local === 'endnoteRef' ||
      local === 'w:endnoteRef' ||
      local === 'footnoteReference' ||
      local === 'w:footnoteReference' ||
      local === 'endnoteReference' ||
      local === 'w:endnoteReference'
    ) {
      return true;
    }
    cur = walker.nextNode() as Element | null;
  }
  return false;
}

/**
 * Checks if a <w:r> element contains an embedded drawing, image, or page field so we never lose it.
 */
function runHasDrawingOrField(rNode: Element): boolean {
  const walker = document.createTreeWalker(rNode, NodeFilter.SHOW_ELEMENT);
  let cur = walker.nextNode() as Element | null;
  while (cur) {
    const local = cur.localName || cur.nodeName;
    if (
      local === 'drawing' ||
      local === 'w:drawing' ||
      local === 'pict' ||
      local === 'w:pict' ||
      local === 'fldChar' ||
      local === 'w:fldChar' ||
      local === 'instrText' ||
      local === 'w:instrText'
    ) {
      return true;
    }
    cur = walker.nextNode() as Element | null;
  }
  return false;
}

interface TokenWithOffset {
  text: string;
  start: number;
  end: number;
}

function tokenizeForStyleAlignment(str: string): TokenWithOffset[] {
  const tokens: TokenWithOffset[] = [];
  const regex = /(\s+|[^\s:;,.|()«»]+|[:;,.|()«»])/g;
  let match: RegExpExecArray | null;
  while ((match = regex.exec(str)) !== null) {
    tokens.push({
      text: match[0],
      start: match.index,
      end: match.index + match[0].length,
    });
  }
  return tokens;
}

/**
 * Maps each token in `newValue` to the exact `<w:rPr>` from `originalText` using Token LCS alignment,
 * preserving inline bold labels, normal values, italics, underlines, font sizes, and colors!
 */
function buildStyledRunsForReplacement(
  originalText: string,
  newValue: string,
  charRPrs: (Element | null)[],
  fallbackRPr: Element | null
): Array<{ text: string; rPr: Element | null }> {
  if (!newValue) return [];

  if (!originalText || charRPrs.length === 0) {
    return [{ text: newValue, rPr: fallbackRPr }];
  }

  const origTokens = tokenizeForStyleAlignment(originalText);
  const newTokens = tokenizeForStyleAlignment(newValue);

  const m = origTokens.length;
  const n = newTokens.length;

  // DP table for Token LCS
  const dp: number[][] = Array.from({ length: m + 1 }, () => new Array(n + 1).fill(0));
  for (let i = m - 1; i >= 0; i--) {
    for (let j = n - 1; j >= 0; j--) {
      if (origTokens[i].text === newTokens[j].text) {
        dp[i][j] = 1 + dp[i + 1][j + 1];
      } else {
        dp[i][j] = Math.max(dp[i + 1][j], dp[i][j + 1]);
      }
    }
  }

  // Backtrack to find matched token pairs (origIdx <-> newIdx)
  const newToOrigMap = new Array<number>(n).fill(-1);
  let i = 0;
  let j = 0;
  while (i < m && j < n) {
    if (origTokens[i].text === newTokens[j].text) {
      newToOrigMap[j] = i;
      i++;
      j++;
    } else if (dp[i + 1][j] >= dp[i][j + 1]) {
      i++;
    } else {
      j++;
    }
  }

  const getOrigTokenRPr = (origTokenIdx: number): Element | null => {
    if (origTokenIdx < 0 || origTokenIdx >= origTokens.length) return fallbackRPr;
    const pos = Math.min(origTokens[origTokenIdx].start, charRPrs.length - 1);
    return pos >= 0 ? charRPrs[pos] : fallbackRPr;
  };

  // Assign an rPr to each token in newTokens
  const tokenRPrs: (Element | null)[] = new Array(n).fill(fallbackRPr);

  for (let idx = 0; idx < n; idx++) {
    const matchedOrig = newToOrigMap[idx];
    if (matchedOrig !== -1) {
      tokenRPrs[idx] = getOrigTokenRPr(matchedOrig);
    } else {
      // Find the nearest matched original token before and after `idx`
      let prevMatchedOrig = -1;
      for (let p = idx - 1; p >= 0; p--) {
        if (newToOrigMap[p] !== -1) {
          prevMatchedOrig = newToOrigMap[p];
          break;
        }
      }
      let nextMatchedOrig = m;
      for (let q = idx + 1; q < n; q++) {
        if (newToOrigMap[q] !== -1) {
          nextMatchedOrig = newToOrigMap[q];
          break;
        }
      }

      // If there were original tokens in the replaced gap (prevMatchedOrig + 1 .. nextMatchedOrig - 1),
      // pick the first non-whitespace token's rPr in that gap (which is the original value's style!)
      let chosenRPr: Element | null = null;
      for (let origGapIdx = prevMatchedOrig + 1; origGapIdx < nextMatchedOrig; origGapIdx++) {
        if (origTokens[origGapIdx].text.trim().length > 0) {
          chosenRPr = getOrigTokenRPr(origGapIdx);
          break;
        }
      }
      if (!chosenRPr && prevMatchedOrig + 1 < nextMatchedOrig) {
        chosenRPr = getOrigTokenRPr(prevMatchedOrig + 1);
      }
      if (!chosenRPr) {
        // Fallback to the last non-bold run in the paragraph if preceded by a colon, or neighbor run
        if (prevMatchedOrig !== -1) {
          chosenRPr = getOrigTokenRPr(Math.min(prevMatchedOrig + 1, m - 1));
        } else if (nextMatchedOrig < m) {
          chosenRPr = getOrigTokenRPr(nextMatchedOrig);
        } else {
          chosenRPr = fallbackRPr;
        }
      }
      tokenRPrs[idx] = chosenRPr;
    }
  }

  // Merge adjacent newTokens that share the exact same rPr node
  const mergedSegments: Array<{ text: string; rPr: Element | null }> = [];
  for (let idx = 0; idx < n; idx++) {
    const tokText = newTokens[idx].text;
    const rPr = tokenRPrs[idx];
    const prev = mergedSegments[mergedSegments.length - 1];
    if (prev && prev.rPr === rPr) {
      prev.text += tokText;
    } else {
      mergedSegments.push({ text: tokText, rPr });
    }
  }

  return mergedSegments;
}

/**
 * Appends text to a `<w:r>` element while properly converting `\n` into `<w:br/>` and `\t` into `<w:tab/>`.
 */
function appendTextWithTabsAndBreaksToRun(
  xmlDoc: Document,
  rElem: Element,
  segmentText: string
): void {
  const lines = segmentText.split('\n');
  lines.forEach((lineText, lineIdx) => {
    if (lineIdx > 0) {
      const brElem = xmlDoc.createElementNS(W_NS, 'w:br');
      rElem.appendChild(brElem);
    }
    const tabParts = lineText.split('\t');
    tabParts.forEach((part, tabIdx) => {
      if (tabIdx > 0) {
        const tabElem = xmlDoc.createElementNS(W_NS, 'w:tab');
        rElem.appendChild(tabElem);
      }
      if (part.length > 0) {
        const tElem = xmlDoc.createElementNS(W_NS, 'w:t');
        tElem.setAttribute('xml:space', 'preserve');
        tElem.textContent = part;
        rElem.appendChild(tElem);
      }
    });
  });
}

function applyReplacementsToTargets(
  targets: XmlParagraphTarget[],
  replacements: Record<number, string>
): Set<string> {
  const modifiedXmlPaths = new Set<string>();

  Object.entries(replacements).forEach(([idxStr, newValue]) => {
    const idx = Number(idxStr);
    if (Number.isNaN(idx) || idx < 0 || idx >= targets.length) return;

    const { xmlDoc, xmlPath, pNode, section } = targets[idx];

    // Extract original text and per-character <w:rPr> map
    const { text: originalText, charRPrs } = extractParagraphTextRunsAndRefs(pNode);

    // CRITICAL: If the replacement text is identical to the original text in the Word XML,
    // do NOT touch the paragraph's XML nodes at all! This keeps 100% of original Word formatting intact.
    if (newValue === originalText) {
      return;
    }

    const pPr = findFirstChildByLocalName(pNode, 'pPr');

    let fallbackStyleRun: Element | null = null;
    const preservedStartRuns: Node[] = [];
    const preservedEndRuns: Node[] = [];

    for (let i = 0; i < pNode.childNodes.length; i++) {
      const child = pNode.childNodes[i] as Element;
      if (child.nodeType === 1 && (child.localName === 'r' || child.nodeName === 'w:r')) {
        if (runHasNoteMarker(child)) {
          const clone = child.cloneNode(true) as Element;
          const tNodes = Array.from(clone.childNodes).filter(
            (n) =>
              n.nodeType === 1 &&
              ((n as Element).localName === 't' || (n as Element).nodeName === 'w:t')
          );
          tNodes.forEach((t) => clone.removeChild(t));
          if (section === 'footnote' || section === 'endnote') {
            preservedStartRuns.push(clone);
          } else {
            preservedEndRuns.push(clone);
          }
        } else if (runHasDrawingOrField(child)) {
          preservedStartRuns.push(child.cloneNode(true));
        } else if (!fallbackStyleRun) {
          fallbackStyleRun = child;
        }
      }
    }

    if (!fallbackStyleRun) {
      fallbackStyleRun = findFirstChildByLocalName(pNode, 'r');
    }
    const fallbackRPr = fallbackStyleRun
      ? findFirstChildByLocalName(fallbackStyleRun, 'rPr')
      : null;

    const normalizedNewText =
      (section === 'footnote' || section === 'endnote') &&
      preservedStartRuns.length > 0 &&
      !newValue.startsWith(' ')
        ? ` ${newValue.trimStart()}`
        : newValue;

    // Build styled run segments using Token LCS alignment so bold prefixes, normal values, etc. are preserved
    const styledSegments = buildStyledRunsForReplacement(
      originalText,
      normalizedNewText,
      charRPrs,
      fallbackRPr
    );

    const childrenToRemove: Node[] = [];
    for (let i = 0; i < pNode.childNodes.length; i++) {
      const child = pNode.childNodes[i];
      if (child !== pPr) {
        childrenToRemove.push(child);
      }
    }
    childrenToRemove.forEach((node) => pNode.removeChild(node));

    preservedStartRuns.forEach((run) => pNode.appendChild(run));

    if (styledSegments.length === 0) {
      const rElem = xmlDoc.createElementNS(W_NS, 'w:r');
      if (fallbackRPr) {
        rElem.appendChild(fallbackRPr.cloneNode(true));
      }
      const tElem = xmlDoc.createElementNS(W_NS, 'w:t');
      tElem.textContent = '';
      rElem.appendChild(tElem);
      pNode.appendChild(rElem);
    } else {
      styledSegments.forEach((seg) => {
        const rElem = xmlDoc.createElementNS(W_NS, 'w:r');
        if (seg.rPr) {
          rElem.appendChild(seg.rPr.cloneNode(true));
        }
        appendTextWithTabsAndBreaksToRun(xmlDoc, rElem, seg.text);
        pNode.appendChild(rElem);
      });
    }

    preservedEndRuns.forEach((run) => pNode.appendChild(run));

    modifiedXmlPaths.add(xmlPath);
  });

  return modifiedXmlPaths;
}

/**
 * Exports a single filled .docx Blob for one record, updating Headers, Body, Tables, Footnotes, Endnotes, and Footers
 * while preserving inline run styles, tabs, line breaks, tables, and note markers.
 */
export async function exportFilledDocxBlob(
  base64Data: string,
  replacements: Record<number, string>
): Promise<Blob> {
  const bytes = base64ToUint8Array(base64Data);
  const zip = await JSZip.loadAsync(bytes);

  const { targets, xmlDocsByPath } = await collectAllDocxParagraphTargets(zip);
  const modifiedXmlPaths = applyReplacementsToTargets(targets, replacements);

  const serializer = new XMLSerializer();
  modifiedXmlPaths.forEach((xmlPath) => {
    const xmlDoc = xmlDocsByPath.get(xmlPath);
    if (xmlDoc) {
      const updatedXml = serializer.serializeToString(xmlDoc);
      zip.file(xmlPath, updatedXml);
    }
  });

  return await zip.generateAsync({
    type: 'blob',
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  });
}

export interface SampleParagraphSpec {
  text: string;
  align?: 'left' | 'center' | 'right' | 'both';
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  fontSizeHalfPt?: number;
  footnoteRefId?: number;
}

export interface SampleDocxOptions {
  headerLines?: SampleParagraphSpec[];
  footerLines?: SampleParagraphSpec[];
  footnotes?: Array<{ id: number; text: string }>;
}

function escapeXml(unsafe: string): string {
  return unsafe
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

/**
 * Generates a complete, valid Microsoft Word (.docx) file in Base64 containing:
 * - Header (word/header1.xml — Headnote / Tiêu đề đầu trang)
 * - Body (word/document.xml — with mixed bold label + normal value runs for realistic Word forms)
 * - Footnotes (word/footnotes.xml — Chú thích chân trang)
 * - Footer (word/footer1.xml — Chân trang)
 */
export async function generateSampleDocxBase64(
  paragraphs: SampleParagraphSpec[],
  options?: SampleDocxOptions
): Promise<string> {
  const zip = new JSZip();

  const headerLines = options?.headerLines || [
    {
      text: 'BIỂU MẪU HÀNH CHÍNH CHUẨN ISO 9001:2015 — MÃ BIỂU MẪU: BM-HCNS-2026',
      align: 'right',
      italic: true,
      fontSizeHalfPt: 20,
    },
  ];

  const footerLines = options?.footerLines || [
    {
      text: 'Tài liệu lưu hành nội bộ · Trang 1/1',
      align: 'center',
      italic: true,
      fontSizeHalfPt: 20,
    },
  ];

  const footnotes = options?.footnotes || [
    {
      id: 1,
      text: 'Chú thích (Footnote 1): Đề nghị đính kèm đầy đủ hóa đơn, chứng từ gốc hoặc văn bản phê duyệt liên quan khi trình ký.',
    },
  ];

  const contentTypesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
  <Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>
  <Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>
</Types>`;

  const rootRelsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`;

  const docRelsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>
  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/>
</Relationships>`;

  const stylesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault>
      <w:rPr>
        <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>
        <w:sz w:val="26"/>
        <w:szCs w:val="26"/>
      </w:rPr>
    </w:rPrDefault>
  </w:docDefaults>
</w:styles>`;

  const buildRunXml = (text: string, sz: number, bold?: boolean, italic?: boolean, underline?: boolean) => {
    if (!text) return '';
    const bTag = bold ? '<w:b/>' : '';
    const iTag = italic ? '<w:i/>' : '';
    const uTag = underline ? '<w:u w:val="single"/>' : '';
    const rPr = `<w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="${sz}"/><w:szCs w:val="${sz}"/>${bTag}${iTag}${uTag}</w:rPr>`;
    return `<w:r>${rPr}<w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r>`;
  };

  const renderParagraphSpecXml = (p: SampleParagraphSpec) => {
    const alignTag = p.align ? `<w:jc w:val="${p.align}"/>` : '';
    const pPr = `<w:pPr>${alignTag}<w:spacing w:after="120" w:line="276" w:lineRule="auto"/></w:pPr>`;
    const sz = p.fontSizeHalfPt || 26;

    let runsXml = '';
    // If paragraph is not marked whole-bold, but has "Label: Value" structure, make the label prefix bold just like standard Vietnamese administrative forms!
    if (!p.bold && p.text.includes(':') && p.align === 'left') {
      const segments = p.text.split(' | ');
      runsXml = segments
        .map((seg, segIdx) => {
          const prefixDivider = segIdx > 0 ? buildRunXml(' | ', sz, false, p.italic, p.underline) : '';
          const colonPos = seg.indexOf(':');
          if (colonPos > 0 && colonPos < 45) {
            const labelPart = seg.slice(0, colonPos + 1);
            const valPart = seg.slice(colonPos + 1);
            return (
              prefixDivider +
              buildRunXml(labelPart, sz, true, p.italic, p.underline) +
              buildRunXml(valPart, sz, false, p.italic, p.underline)
            );
          }
          return prefixDivider + buildRunXml(seg, sz, p.bold, p.italic, p.underline);
        })
        .join('');
    } else {
      runsXml = buildRunXml(p.text, sz, p.bold, p.italic, p.underline);
    }

    const fnRef = p.footnoteRefId
      ? `<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="${p.footnoteRefId}"/></w:r>`
      : '';
    return `<w:p>${pPr}${runsXml}${fnRef}</w:p>`;
  };

  const header1Xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  ${headerLines.map(renderParagraphSpecXml).join('\n  ')}
</w:hdr>`;

  const footer1Xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  ${footerLines.map(renderParagraphSpecXml).join('\n  ')}
</w:ftr>`;

  const footnotesItemsXml = footnotes
    .map(
      (fn) => `  <w:footnote w:id="${fn.id}">
    <w:p>
      <w:pPr><w:spacing w:after="60" w:line="240" w:lineRule="auto"/></w:pPr>
      <w:r><w:rPr><w:vertAlign w:val="superscript"/><w:sz w:val="20"/></w:rPr><w:footnoteRef/></w:r>
      <w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="21"/><w:i/></w:rPr><w:t xml:space="preserve"> ${escapeXml(fn.text)}</w:t></w:r>
    </w:p>
  </w:footnote>`
    )
    .join('\n');

  const footnotesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:footnote w:type="separator" w:id="-1">
    <w:p><w:r><w:separator/></w:r></w:p>
  </w:footnote>
  <w:footnote w:type="continuationSeparator" w:id="0">
    <w:p><w:r><w:continuationSeparator/></w:r></w:p>
  </w:footnote>
${footnotesItemsXml}
</w:footnotes>`;

  const bodyParagraphsXml = paragraphs.map(renderParagraphSpecXml).join('\n');

  const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    ${bodyParagraphsXml}
    <w:sectPr>
      <w:headerReference w:type="default" r:id="rId2"/>
      <w:footerReference w:type="default" r:id="rId3"/>
      <w:pgSz w:w="11906" w:h="16838"/>
      <w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1701" w:header="708" w:footer="708" w:gutter="0"/>
    </w:sectPr>
  </w:body>
</w:document>`;

  zip.file('[Content_Types].xml', contentTypesXml);
  zip.file('_rels/.rels', rootRelsXml);
  zip.file('word/_rels/document.xml.rels', docRelsXml);
  zip.file('word/styles.xml', stylesXml);
  zip.file('word/header1.xml', header1Xml);
  zip.file('word/footer1.xml', footer1Xml);
  zip.file('word/footnotes.xml', footnotesXml);
  zip.file('word/document.xml', documentXml);

  return await zip.generateAsync({ type: 'base64' });
}

export function triggerBlobDownload(blob: Blob, fileName: string): void {
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = fileName;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  setTimeout(() => URL.revokeObjectURL(url), 2500);
}
