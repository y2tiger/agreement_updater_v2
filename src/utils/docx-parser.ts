/**
 * DOCX Parser Utilities
 * Handles parsing and extracting structured data from DOCX files
 */

import AdmZip from 'adm-zip';
import { parseStringPromise, Builder } from 'xml2js';
import {
  DocxParagraph,
  DocxRun,
  DocxTable,
  DocxTableRow,
  DocxTableCell,
  DocxHeaderFooter,
  RunFormatting,
  NumberingInfo,
} from '../types';

// XML namespace prefixes used in DOCX
const NAMESPACES = {
  w: 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
  r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
};

export interface ParsedDocx {
  zip: AdmZip;
  documentXml: unknown;
  documentPath: string;
  paragraphs: DocxParagraph[];
  tables: DocxTable[];
  headersFooters: DocxHeaderFooter[];
  numbering?: unknown;
  styles?: unknown;
  relationships?: unknown;
}

/**
 * Parse a DOCX buffer into structured data
 */
export async function parseDocx(buffer: Buffer): Promise<ParsedDocx> {
  const zip = new AdmZip(buffer);

  // Read the main document
  const documentEntry = zip.getEntry('word/document.xml');
  if (!documentEntry) {
    throw new Error('Invalid DOCX: word/document.xml not found');
  }

  const documentXmlStr = documentEntry.getData().toString('utf-8');
  const documentXml = await parseStringPromise(documentXmlStr, {
    explicitArray: true,
    preserveChildrenOrder: true,
  });

  // Parse paragraphs and tables from the document body
  const body = getBody(documentXml);
  const { paragraphs, tables } = parseBodyElements(body);

  // Parse headers and footers
  const headersFooters = await parseHeadersFooters(zip);

  // Parse numbering definitions if present
  let numbering: unknown;
  const numberingEntry = zip.getEntry('word/numbering.xml');
  if (numberingEntry) {
    numbering = await parseStringPromise(numberingEntry.getData().toString('utf-8'), {
      explicitArray: true,
    });
  }

  // Parse styles if present
  let styles: unknown;
  const stylesEntry = zip.getEntry('word/styles.xml');
  if (stylesEntry) {
    styles = await parseStringPromise(stylesEntry.getData().toString('utf-8'), {
      explicitArray: true,
    });
  }

  // Parse relationships
  let relationships: unknown;
  const relsEntry = zip.getEntry('word/_rels/document.xml.rels');
  if (relsEntry) {
    relationships = await parseStringPromise(relsEntry.getData().toString('utf-8'), {
      explicitArray: true,
    });
  }

  return {
    zip,
    documentXml,
    documentPath: 'word/document.xml',
    paragraphs,
    tables,
    headersFooters,
    numbering,
    styles,
    relationships,
  };
}

/**
 * Find a key in an object that matches a pattern (handles namespace variations)
 */
function findKey(obj: Record<string, unknown>, patterns: string[]): string | undefined {
  const keys = Object.keys(obj);
  for (const pattern of patterns) {
    const found = keys.find(k => k === pattern || k.endsWith(':' + pattern.split(':').pop()));
    if (found) return found;
  }
  return undefined;
}

/**
 * Get the document body from parsed XML
 */
function getBody(documentXml: unknown): unknown {
  const doc = documentXml as Record<string, unknown>;

  // Try different possible document element names
  const docKey = findKey(doc, ['w:document', 'document']);
  if (!docKey) {
    // Debug: log available keys
    console.error('Available root keys:', Object.keys(doc));
    throw new Error('Invalid DOCX structure: w:document not found');
  }

  const document = doc[docKey] as Record<string, unknown> | Record<string, unknown>[];
  const docElement = Array.isArray(document) ? document[0] : document;

  if (!docElement) {
    throw new Error('Invalid DOCX structure: document element is empty');
  }

  // Try different possible body element names
  const bodyKey = findKey(docElement as Record<string, unknown>, ['w:body', 'body']);
  if (!bodyKey) {
    console.error('Available document keys:', Object.keys(docElement as Record<string, unknown>));
    throw new Error('Invalid DOCX structure: w:body not found');
  }

  const body = (docElement as Record<string, unknown>)[bodyKey];
  const bodyElement = Array.isArray(body) ? body[0] : body;

  if (!bodyElement) {
    throw new Error('Invalid DOCX structure: body element is empty');
  }

  return bodyElement;
}

/**
 * Parse body elements (paragraphs and tables)
 */
function parseBodyElements(body: unknown): {
  paragraphs: DocxParagraph[];
  tables: DocxTable[];
} {
  const paragraphs: DocxParagraph[] = [];
  const tables: DocxTable[] = [];

  const bodyObj = body as Record<string, unknown[]>;
  let paragraphIndex = 0;
  let tableIndex = 0;

  // Process paragraphs
  const pElements = bodyObj['w:p'] || [];
  for (const pElement of pElements) {
    const paragraph = parseParagraph(pElement, paragraphIndex);
    paragraphs.push(paragraph);
    paragraphIndex++;
  }

  // Process tables
  const tblElements = bodyObj['w:tbl'] || [];
  for (const tblElement of tblElements) {
    const table = parseTable(tblElement, tableIndex);
    tables.push(table);
    tableIndex++;
  }

  return { paragraphs, tables };
}

/**
 * Parse a paragraph element
 */
function parseParagraph(pElement: unknown, index: number): DocxParagraph {
  const pObj = pElement as Record<string, unknown[]>;
  const runs: DocxRun[] = [];
  let currentOffset = 0;
  let runIndex = 0;

  // Extract style from paragraph properties
  let style: string | undefined;
  let numbering: NumberingInfo | undefined;

  const pPr = pObj['w:pPr'];
  if (pPr && pPr[0]) {
    const pPrObj = pPr[0] as Record<string, unknown[]>;

    // Get style
    const pStyle = pPrObj['w:pStyle'];
    if (pStyle && pStyle[0]) {
      const styleObj = pStyle[0] as Record<string, unknown>;
      const attrs = styleObj['$'] as Record<string, string>;
      if (attrs && attrs['w:val']) {
        style = attrs['w:val'];
      }
    }

    // Get numbering
    const numPr = pPrObj['w:numPr'];
    if (numPr && numPr[0]) {
      const numPrObj = numPr[0] as Record<string, unknown[]>;
      const ilvl = numPrObj['w:ilvl'];
      const numId = numPrObj['w:numId'];

      if (ilvl && numId) {
        const ilvlAttrs = (ilvl[0] as Record<string, unknown>)['$'] as Record<string, string>;
        const numIdAttrs = (numId[0] as Record<string, unknown>)['$'] as Record<string, string>;

        if (ilvlAttrs && numIdAttrs) {
          numbering = {
            numId: numIdAttrs['w:val'] || '',
            level: parseInt(ilvlAttrs['w:val'] || '0', 10),
            format: '', // Would need to look up in numbering.xml
          };
        }
      }
    }
  }

  // Process runs
  const rElements = pObj['w:r'] || [];
  for (const rElement of rElements) {
    const run = parseRun(rElement, runIndex, currentOffset);
    runs.push(run);
    currentOffset += run.text.length;
    runIndex++;
  }

  // Build full text and normalized text
  const text = runs.map(r => r.text).join('');
  const normalizedText = normalizeText(text);

  return {
    index,
    text,
    normalizedText,
    runs,
    style,
    numbering,
    xmlElement: pElement,
  };
}

/**
 * Parse a run element
 */
function parseRun(rElement: unknown, index: number, startOffset: number): DocxRun {
  const rObj = rElement as Record<string, unknown[]>;
  let text = '';

  // Extract text from w:t elements
  const tElements = rObj['w:t'] || [];
  for (const tElement of tElements) {
    if (typeof tElement === 'string') {
      text += tElement;
    } else if (typeof tElement === 'object' && tElement !== null) {
      const tObj = tElement as Record<string, unknown>;
      if (tObj['_']) {
        text += tObj['_'] as string;
      }
    }
  }

  // Handle special elements (tabs, breaks, etc.)
  if (rObj['w:tab']) {
    text += '\t';
  }
  if (rObj['w:br']) {
    text += '\n';
  }

  // Extract formatting
  const formatting = parseRunFormatting(rObj['w:rPr']);

  return {
    index,
    text,
    startOffset,
    endOffset: startOffset + text.length,
    formatting,
    xmlElement: rElement,
  };
}

/**
 * Parse run formatting properties
 */
function parseRunFormatting(rPr: unknown[] | undefined): RunFormatting {
  const formatting: RunFormatting = {};

  if (!rPr || !rPr[0]) {
    return formatting;
  }

  const rPrObj = rPr[0] as Record<string, unknown[]>;

  // Bold
  if (rPrObj['w:b']) {
    formatting.bold = true;
  }

  // Italic
  if (rPrObj['w:i']) {
    formatting.italic = true;
  }

  // Underline
  if (rPrObj['w:u']) {
    formatting.underline = true;
  }

  // Highlight
  const highlight = rPrObj['w:highlight'];
  if (highlight && highlight[0]) {
    const hlObj = highlight[0] as Record<string, unknown>;
    const attrs = hlObj['$'] as Record<string, string>;
    if (attrs && attrs['w:val']) {
      formatting.highlight = attrs['w:val'];
    }
  }

  // Font size
  const sz = rPrObj['w:sz'];
  if (sz && sz[0]) {
    const szObj = sz[0] as Record<string, unknown>;
    const attrs = szObj['$'] as Record<string, string>;
    if (attrs && attrs['w:val']) {
      formatting.fontSize = parseInt(attrs['w:val'], 10) / 2; // Half-points to points
    }
  }

  // Font name
  const rFonts = rPrObj['w:rFonts'];
  if (rFonts && rFonts[0]) {
    const fontsObj = rFonts[0] as Record<string, unknown>;
    const attrs = fontsObj['$'] as Record<string, string>;
    if (attrs && (attrs['w:ascii'] || attrs['w:hAnsi'])) {
      formatting.fontName = attrs['w:ascii'] || attrs['w:hAnsi'];
    }
  }

  return formatting;
}

/**
 * Parse a table element
 */
function parseTable(tblElement: unknown, index: number): DocxTable {
  const tblObj = tblElement as Record<string, unknown[]>;
  const rows: DocxTableRow[] = [];

  const trElements = tblObj['w:tr'] || [];
  let rowIndex = 0;

  for (const trElement of trElements) {
    const row = parseTableRow(trElement, rowIndex);
    rows.push(row);
    rowIndex++;
  }

  return {
    index,
    rows,
    xmlElement: tblElement,
  };
}

/**
 * Parse a table row element
 */
function parseTableRow(trElement: unknown, rowIndex: number): DocxTableRow {
  const trObj = trElement as Record<string, unknown[]>;
  const cells: DocxTableCell[] = [];

  const tcElements = trObj['w:tc'] || [];
  let cellIndex = 0;

  for (const tcElement of tcElements) {
    const cell = parseTableCell(tcElement, rowIndex, cellIndex);
    cells.push(cell);
    cellIndex++;
  }

  return {
    index: rowIndex,
    cells,
    xmlElement: trElement,
  };
}

/**
 * Parse a table cell element
 */
function parseTableCell(tcElement: unknown, rowIndex: number, cellIndex: number): DocxTableCell {
  const tcObj = tcElement as Record<string, unknown[]>;
  const paragraphs: DocxParagraph[] = [];

  const pElements = tcObj['w:p'] || [];
  let paragraphIndex = 0;

  for (const pElement of pElements) {
    const paragraph = parseParagraph(pElement, paragraphIndex);
    paragraphs.push(paragraph);
    paragraphIndex++;
  }

  return {
    rowIndex,
    cellIndex,
    paragraphs,
    xmlElement: tcElement,
  };
}

/**
 * Parse headers and footers from the DOCX
 */
async function parseHeadersFooters(zip: AdmZip): Promise<DocxHeaderFooter[]> {
  const headersFooters: DocxHeaderFooter[] = [];
  const entries = zip.getEntries();

  for (const entry of entries) {
    if (entry.entryName.match(/^word\/(header|footer)\d+\.xml$/)) {
      const type = entry.entryName.includes('header') ? 'header' : 'footer';
      const match = entry.entryName.match(/(\d+)/);
      const sectionIndex = match ? parseInt(match[1], 10) - 1 : 0;

      const xmlStr = entry.getData().toString('utf-8');
      const xml = await parseStringPromise(xmlStr, {
        explicitArray: true,
        preserveChildrenOrder: true,
      });

      // Get the root element (w:hdr or w:ftr)
      const rootKey = type === 'header' ? 'w:hdr' : 'w:ftr';
      const root = xml[rootKey];

      if (root && root[0]) {
        const { paragraphs } = parseBodyElements(root[0]);

        headersFooters.push({
          type,
          sectionIndex,
          paragraphs,
          xmlPath: entry.entryName,
        });
      }
    }
  }

  return headersFooters;
}

/**
 * Normalize text for matching (lowercase, collapse whitespace)
 */
export function normalizeText(text: string): string {
  return text
    .toLowerCase()
    .replace(/\s+/g, ' ')
    .trim();
}

/**
 * Extract plain text from a parsed DOCX for LLM analysis
 */
export function extractPlainText(parsed: ParsedDocx): string {
  const lines: string[] = [];

  for (const paragraph of parsed.paragraphs) {
    if (paragraph.text.trim()) {
      lines.push(paragraph.text);
    }
  }

  // Add table content
  for (const table of parsed.tables) {
    lines.push('\n[TABLE]');
    for (const row of table.rows) {
      const cellTexts = row.cells.map(cell =>
        cell.paragraphs.map(p => p.text).join(' ')
      );
      lines.push('| ' + cellTexts.join(' | ') + ' |');
    }
    lines.push('[/TABLE]\n');
  }

  // Add header/footer content
  for (const hf of parsed.headersFooters) {
    const typeLabel = hf.type.toUpperCase();
    lines.push(`\n[${typeLabel}]`);
    for (const paragraph of hf.paragraphs) {
      if (paragraph.text.trim()) {
        lines.push(paragraph.text);
      }
    }
    lines.push(`[/${typeLabel}]\n`);
  }

  return lines.join('\n');
}

/**
 * Serialize modified XML back to the DOCX
 */
export function serializeXml(xmlObj: unknown): string {
  const builder = new Builder({
    renderOpts: { pretty: false },
    xmldec: { version: '1.0', encoding: 'UTF-8', standalone: true },
  });
  return builder.buildObject(xmlObj);
}

/**
 * Save the modified DOCX to a buffer
 */
export function saveDocx(parsed: ParsedDocx, documentXml: unknown): Buffer {
  const zip = parsed.zip;

  // Serialize and update the main document
  const xmlStr = serializeXml(documentXml);
  zip.updateFile('word/document.xml', Buffer.from(xmlStr, 'utf-8'));

  return zip.toBuffer();
}

/**
 * Find paragraph by normalized text match
 */
export function findParagraphByText(
  paragraphs: DocxParagraph[],
  searchText: string,
  fuzzy: boolean = false
): DocxParagraph | undefined {
  const normalizedSearch = normalizeText(searchText);

  // Exact match first
  for (const paragraph of paragraphs) {
    if (paragraph.normalizedText === normalizedSearch) {
      return paragraph;
    }
  }

  // Contains match
  for (const paragraph of paragraphs) {
    if (paragraph.normalizedText.includes(normalizedSearch)) {
      return paragraph;
    }
  }

  // Fuzzy match if enabled
  if (fuzzy) {
    for (const paragraph of paragraphs) {
      if (fuzzyMatch(paragraph.normalizedText, normalizedSearch)) {
        return paragraph;
      }
    }
  }

  return undefined;
}

/**
 * Simple fuzzy matching (allows for minor differences)
 */
function fuzzyMatch(text: string, search: string): boolean {
  // Remove all punctuation and compare
  const cleanText = text.replace(/[^\w\s]/g, '');
  const cleanSearch = search.replace(/[^\w\s]/g, '');

  return cleanText.includes(cleanSearch);
}
