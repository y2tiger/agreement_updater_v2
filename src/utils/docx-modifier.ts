/**
 * DOCX Modifier Utilities
 * Handles safe modifications to DOCX structure
 */

import { parseStringPromise, Builder } from 'xml2js';
import {
  ParsedDocx,
  DocxParagraph,
  ElementLocation,
  EditSpec,
  ContractEditError,
} from '../types';
import { normalizeText } from './docx-parser';

/**
 * Apply an edit spec to the document
 */
export async function applyEdit(
  parsed: ParsedDocx,
  editSpec: EditSpec,
  location: ElementLocation
): Promise<{ success: boolean; modifiedXml: unknown; error?: string }> {
  try {
    // Deep clone the document XML to avoid mutations
    const documentXml = JSON.parse(JSON.stringify(parsed.documentXml));

    switch (location.type) {
      case 'paragraph':
      case 'run':
        return await applyParagraphEdit(documentXml, editSpec, location);
      case 'table_cell':
        return await applyTableCellEdit(documentXml, editSpec, location);
      case 'header':
      case 'footer':
        // Headers/footers need special handling
        return { success: false, modifiedXml: documentXml, error: 'Header/footer edits not yet implemented' };
      default:
        return { success: false, modifiedXml: documentXml, error: `Unknown location type: ${location.type}` };
    }
  } catch (error) {
    return {
      success: false,
      modifiedXml: parsed.documentXml,
      error: error instanceof Error ? error.message : 'Unknown error during edit',
    };
  }
}

/**
 * Apply an edit to a paragraph
 */
async function applyParagraphEdit(
  documentXml: unknown,
  editSpec: EditSpec,
  location: ElementLocation
): Promise<{ success: boolean; modifiedXml: unknown; error?: string }> {
  const body = getBodyFromXml(documentXml);
  const paragraphs = body['w:p'] as unknown[];

  if (location.paragraphIndex === undefined || location.paragraphIndex >= paragraphs.length) {
    return { success: false, modifiedXml: documentXml, error: 'Invalid paragraph index' };
  }

  const paragraph = paragraphs[location.paragraphIndex] as Record<string, unknown[]>;

  switch (editSpec.editType) {
    case 'replace':
      return replaceInParagraph(documentXml, paragraph, editSpec, location);
    case 'insert_before':
      return insertBeforeParagraph(documentXml, paragraphs, location.paragraphIndex, editSpec);
    case 'insert_after':
      return insertAfterParagraph(documentXml, paragraphs, location.paragraphIndex, editSpec);
    case 'delete':
      return deleteParagraph(documentXml, paragraphs, location.paragraphIndex);
    default:
      return { success: false, modifiedXml: documentXml, error: `Unknown edit type: ${editSpec.editType}` };
  }
}

/**
 * Replace text within a paragraph
 */
function replaceInParagraph(
  documentXml: unknown,
  paragraph: Record<string, unknown[]>,
  editSpec: EditSpec,
  location: ElementLocation
): { success: boolean; modifiedXml: unknown; error?: string } {
  const runs = paragraph['w:r'] || [];

  if (editSpec.boundaries.fullElement) {
    // Replace the entire paragraph content
    return replaceEntireParagraph(documentXml, paragraph, editSpec);
  }

  // Find the runs that contain the anchor text
  const { startRunIndex, endRunIndex, startCharOffset, endCharOffset } = findRunsForText(
    runs,
    editSpec.anchorText,
    location.startCharOffset,
    location.endCharOffset
  );

  if (startRunIndex === -1) {
    return { success: false, modifiedXml: documentXml, error: 'Could not locate anchor text in runs' };
  }

  // Modify the runs
  const newRuns = modifyRuns(
    runs,
    startRunIndex,
    endRunIndex,
    startCharOffset,
    endCharOffset,
    editSpec.afterText,
    editSpec.highlight
  );

  paragraph['w:r'] = newRuns;

  return { success: true, modifiedXml: documentXml };
}

/**
 * Replace the entire content of a paragraph
 */
function replaceEntireParagraph(
  documentXml: unknown,
  paragraph: Record<string, unknown[]>,
  editSpec: EditSpec
): { success: boolean; modifiedXml: unknown; error?: string } {
  // Preserve paragraph properties
  const pPr = paragraph['w:pPr'];

  // Create a new run with the replacement text and highlight
  const newRun = createHighlightedRun(editSpec.afterText, editSpec.highlight);

  // Replace all runs with the new run
  paragraph['w:r'] = [newRun];

  // Restore paragraph properties
  if (pPr) {
    paragraph['w:pPr'] = pPr;
  }

  return { success: true, modifiedXml: documentXml };
}

/**
 * Find which runs contain the target text
 */
function findRunsForText(
  runs: unknown[],
  anchorText: string,
  expectedStartOffset: number,
  expectedEndOffset: number
): { startRunIndex: number; endRunIndex: number; startCharOffset: number; endCharOffset: number } {
  const normalizedAnchor = normalizeText(anchorText);
  let currentOffset = 0;
  let runTexts: { index: number; text: string; start: number; end: number }[] = [];

  // Build a map of run positions
  for (let i = 0; i < runs.length; i++) {
    const run = runs[i] as Record<string, unknown[]>;
    const text = extractRunText(run);
    runTexts.push({
      index: i,
      text,
      start: currentOffset,
      end: currentOffset + text.length,
    });
    currentOffset += text.length;
  }

  // Concatenate all run text
  const fullText = runTexts.map(r => r.text).join('');
  const normalizedFull = normalizeText(fullText);

  // Find the anchor in the normalized text
  const anchorPos = normalizedFull.indexOf(normalizedAnchor);
  if (anchorPos === -1) {
    return { startRunIndex: -1, endRunIndex: -1, startCharOffset: -1, endCharOffset: -1 };
  }

  // Map back to original positions (accounting for whitespace normalization)
  const startOffset = mapNormalizedToOriginal(fullText, anchorPos);
  const endOffset = mapNormalizedToOriginal(fullText, anchorPos + normalizedAnchor.length);

  // Find which runs contain these offsets
  let startRunIndex = -1;
  let endRunIndex = -1;
  let startCharOffset = 0;
  let endCharOffset = 0;

  for (const runInfo of runTexts) {
    if (startRunIndex === -1 && runInfo.end > startOffset) {
      startRunIndex = runInfo.index;
      startCharOffset = startOffset - runInfo.start;
    }
    if (runInfo.end >= endOffset) {
      endRunIndex = runInfo.index;
      endCharOffset = endOffset - runInfo.start;
      break;
    }
  }

  return { startRunIndex, endRunIndex, startCharOffset, endCharOffset };
}

/**
 * Map a position in normalized text back to original text
 */
function mapNormalizedToOriginal(original: string, normalizedPos: number): number {
  const normalized = normalizeText(original);
  let origIndex = 0;
  let normIndex = 0;

  while (normIndex < normalizedPos && origIndex < original.length) {
    const origChar = original[origIndex];
    const normChar = normalized[normIndex];

    if (origChar.toLowerCase() === normChar || (origChar.match(/\s/) && normChar === ' ')) {
      normIndex++;
    }
    origIndex++;
  }

  return origIndex;
}

/**
 * Extract text from a run element
 */
function extractRunText(run: Record<string, unknown[]>): string {
  let text = '';

  const tElements = run['w:t'] || [];
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

  if (run['w:tab']) text += '\t';
  if (run['w:br']) text += '\n';

  return text;
}

/**
 * Modify runs to apply the replacement
 */
function modifyRuns(
  runs: unknown[],
  startRunIndex: number,
  endRunIndex: number,
  startCharOffset: number,
  endCharOffset: number,
  newText: string,
  highlight: boolean
): unknown[] {
  const result: unknown[] = [];

  for (let i = 0; i < runs.length; i++) {
    if (i < startRunIndex || i > endRunIndex) {
      // Keep unchanged runs
      result.push(runs[i]);
    } else if (i === startRunIndex && i === endRunIndex) {
      // Single run contains entire replacement
      const run = runs[i] as Record<string, unknown[]>;
      const text = extractRunText(run);

      // Text before the replacement
      if (startCharOffset > 0) {
        const beforeRun = cloneRunWithText(run, text.substring(0, startCharOffset));
        result.push(beforeRun);
      }

      // The replacement text with highlight
      const replacementRun = createHighlightedRun(newText, highlight, run);
      result.push(replacementRun);

      // Text after the replacement
      if (endCharOffset < text.length) {
        const afterRun = cloneRunWithText(run, text.substring(endCharOffset));
        result.push(afterRun);
      }
    } else if (i === startRunIndex) {
      // Start of multi-run replacement
      const run = runs[i] as Record<string, unknown[]>;
      const text = extractRunText(run);

      if (startCharOffset > 0) {
        const beforeRun = cloneRunWithText(run, text.substring(0, startCharOffset));
        result.push(beforeRun);
      }

      // Add the replacement text at the start run position
      const replacementRun = createHighlightedRun(newText, highlight, run);
      result.push(replacementRun);
    } else if (i === endRunIndex) {
      // End of multi-run replacement
      const run = runs[i] as Record<string, unknown[]>;
      const text = extractRunText(run);

      if (endCharOffset < text.length) {
        const afterRun = cloneRunWithText(run, text.substring(endCharOffset));
        result.push(afterRun);
      }
    }
    // Runs between start and end are removed (replaced by the new text)
  }

  return result;
}

/**
 * Clone a run with new text
 */
function cloneRunWithText(originalRun: Record<string, unknown[]>, newText: string): Record<string, unknown[]> {
  const run: Record<string, unknown[]> = {};

  // Copy run properties
  if (originalRun['w:rPr']) {
    run['w:rPr'] = JSON.parse(JSON.stringify(originalRun['w:rPr']));
  }

  // Set new text
  run['w:t'] = [{ _: newText, $: { 'xml:space': 'preserve' } }];

  return run;
}

/**
 * Create a run with highlighted text
 */
function createHighlightedRun(
  text: string,
  highlight: boolean,
  templateRun?: Record<string, unknown[]>
): Record<string, unknown[]> {
  const run: Record<string, unknown[]> = {};

  // Copy existing run properties or create new ones
  if (templateRun && templateRun['w:rPr']) {
    run['w:rPr'] = JSON.parse(JSON.stringify(templateRun['w:rPr']));
  } else {
    run['w:rPr'] = [{}];
  }

  // Add yellow highlight if requested
  if (highlight) {
    const rPr = run['w:rPr'][0] as Record<string, unknown[]>;
    rPr['w:highlight'] = [{ $: { 'w:val': 'yellow' } }];
  }

  // Set the text
  run['w:t'] = [{ _: text, $: { 'xml:space': 'preserve' } }];

  return run;
}

/**
 * Insert a new paragraph before the target
 */
function insertBeforeParagraph(
  documentXml: unknown,
  paragraphs: unknown[],
  targetIndex: number,
  editSpec: EditSpec
): { success: boolean; modifiedXml: unknown; error?: string } {
  const newParagraph = createParagraph(editSpec.afterText, editSpec.highlight);
  paragraphs.splice(targetIndex, 0, newParagraph);
  return { success: true, modifiedXml: documentXml };
}

/**
 * Insert a new paragraph after the target
 */
function insertAfterParagraph(
  documentXml: unknown,
  paragraphs: unknown[],
  targetIndex: number,
  editSpec: EditSpec
): { success: boolean; modifiedXml: unknown; error?: string } {
  const newParagraph = createParagraph(editSpec.afterText, editSpec.highlight);
  paragraphs.splice(targetIndex + 1, 0, newParagraph);
  return { success: true, modifiedXml: documentXml };
}

/**
 * Delete a paragraph
 */
function deleteParagraph(
  documentXml: unknown,
  paragraphs: unknown[],
  targetIndex: number
): { success: boolean; modifiedXml: unknown; error?: string } {
  paragraphs.splice(targetIndex, 1);
  return { success: true, modifiedXml: documentXml };
}

/**
 * Create a new paragraph element
 */
function createParagraph(text: string, highlight: boolean): Record<string, unknown[]> {
  const run = createHighlightedRun(text, highlight);
  return {
    'w:r': [run],
  };
}

/**
 * Apply an edit to a table cell
 */
async function applyTableCellEdit(
  documentXml: unknown,
  editSpec: EditSpec,
  location: ElementLocation
): Promise<{ success: boolean; modifiedXml: unknown; error?: string }> {
  const body = getBodyFromXml(documentXml);
  const tables = body['w:tbl'] as unknown[];

  if (location.tableIndex === undefined || location.tableIndex >= tables.length) {
    return { success: false, modifiedXml: documentXml, error: 'Invalid table index' };
  }

  const table = tables[location.tableIndex] as Record<string, unknown[]>;
  const rows = table['w:tr'] || [];

  if (location.rowIndex === undefined || location.rowIndex >= rows.length) {
    return { success: false, modifiedXml: documentXml, error: 'Invalid row index' };
  }

  const row = rows[location.rowIndex] as Record<string, unknown[]>;
  const cells = row['w:tc'] || [];

  if (location.cellIndex === undefined || location.cellIndex >= cells.length) {
    return { success: false, modifiedXml: documentXml, error: 'Invalid cell index' };
  }

  const cell = cells[location.cellIndex] as Record<string, unknown[]>;
  const cellParagraphs = cell['w:p'] || [];

  // For simplicity, apply the edit to the first paragraph in the cell
  if (cellParagraphs.length === 0) {
    return { success: false, modifiedXml: documentXml, error: 'No paragraphs in cell' };
  }

  const paragraph = cellParagraphs[0] as Record<string, unknown[]>;
  return replaceInParagraph(documentXml, paragraph, editSpec, {
    ...location,
    type: 'paragraph',
  });
}

/**
 * Get the body element from document XML
 */
function getBodyFromXml(documentXml: unknown): Record<string, unknown[]> {
  const doc = documentXml as Record<string, unknown>;
  const document = doc['w:document'] as unknown[];
  const documentObj = document[0] as Record<string, unknown>;
  const body = documentObj['w:body'] as unknown[];
  return body[0] as Record<string, unknown[]>;
}

/**
 * Verify that the modified document doesn't contain "[object Object]"
 */
export function checkForObjectStringification(documentXml: unknown): boolean {
  const xmlString = JSON.stringify(documentXml);
  return xmlString.includes('[object Object]');
}

/**
 * Verify document integrity after modifications
 */
export function verifyDocumentIntegrity(documentXml: unknown): { valid: boolean; error?: string } {
  try {
    // Check basic structure
    const doc = documentXml as Record<string, unknown>;
    if (!doc['w:document']) {
      return { valid: false, error: 'Missing w:document root element' };
    }

    const document = doc['w:document'] as unknown[];
    if (!document[0]) {
      return { valid: false, error: 'Empty w:document element' };
    }

    const documentObj = document[0] as Record<string, unknown>;
    if (!documentObj['w:body']) {
      return { valid: false, error: 'Missing w:body element' };
    }

    // Check for object stringification
    if (checkForObjectStringification(documentXml)) {
      return { valid: false, error: '[object Object] contamination detected' };
    }

    // Verify XML can be serialized
    const builder = new Builder();
    builder.buildObject(documentXml);

    return { valid: true };
  } catch (error) {
    return {
      valid: false,
      error: error instanceof Error ? error.message : 'Unknown integrity error',
    };
  }
}
