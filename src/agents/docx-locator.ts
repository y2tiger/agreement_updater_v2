/**
 * DOCX Locator Agent
 *
 * Maps Edit Spec targets to actual DOCX elements:
 * - paragraph
 * - run
 * - table cell
 * - header/footer
 *
 * Handles run fragmentation via normalized-text mapping.
 * Must find an exact anchor match or FAIL.
 */

import {
  EditSpec,
  ElementLocation,
  LocatorResult,
  DocxParagraph,
  DocxTable,
  DocxHeaderFooter,
  ContractEditError,
  BlockedError,
} from '../types';
import { ParsedDocx, normalizeText } from '../utils/docx-parser';

interface LocatorContext {
  parsedDocx: ParsedDocx;
  logs: string[];
}

/**
 * DOCX Locator Agent
 */
export class DocxLocatorAgent {
  private logs: string[] = [];

  private log(message: string): void {
    const timestamp = new Date().toISOString();
    this.logs.push(`[${timestamp}] ${message}`);
  }

  /**
   * Locate the target element for an edit spec
   */
  async locate(
    parsedDocx: ParsedDocx,
    editSpec: EditSpec
  ): Promise<LocatorResult> {
    this.logs = [];
    this.log(`Locating target for change ${editSpec.changeId}`);
    this.log(`Anchor text: "${editSpec.anchorText.substring(0, 100)}..."`);

    const normalizedAnchor = normalizeText(editSpec.anchorText);

    // Try to find in paragraphs first
    const paragraphResult = this.findInParagraphs(
      parsedDocx.paragraphs,
      normalizedAnchor,
      editSpec.anchorText
    );

    if (paragraphResult.found) {
      this.log(`Found in paragraph ${paragraphResult.location?.paragraphIndex}`);
      return paragraphResult;
    }

    // Try tables
    const tableResult = this.findInTables(
      parsedDocx.tables,
      normalizedAnchor,
      editSpec.anchorText
    );

    if (tableResult.found) {
      this.log(`Found in table ${tableResult.location?.tableIndex}`);
      return tableResult;
    }

    // Try headers/footers
    const headerFooterResult = this.findInHeadersFooters(
      parsedDocx.headersFooters,
      normalizedAnchor,
      editSpec.anchorText
    );

    if (headerFooterResult.found) {
      this.log(`Found in ${headerFooterResult.location?.type}`);
      return headerFooterResult;
    }

    // Not found
    this.log('Target not found in document');
    return {
      found: false,
      confidence: 0,
      error: 'Anchor text not found in document',
    };
  }

  /**
   * Find anchor text in paragraphs
   */
  private findInParagraphs(
    paragraphs: DocxParagraph[],
    normalizedAnchor: string,
    originalAnchor: string
  ): LocatorResult {
    const matches: ElementLocation[] = [];

    for (const paragraph of paragraphs) {
      // Check for exact match in normalized text
      const normalizedParagraph = paragraph.normalizedText;
      const matchIndex = normalizedParagraph.indexOf(normalizedAnchor);

      if (matchIndex !== -1) {
        // Found a match - now find the exact character positions
        const { startOffset, endOffset } = this.mapNormalizedToOriginal(
          paragraph.text,
          normalizedParagraph,
          matchIndex,
          normalizedAnchor.length
        );

        // Find which runs contain this text
        const runIndices = this.findRunsContainingRange(
          paragraph.runs,
          startOffset,
          endOffset
        );

        const location: ElementLocation = {
          type: runIndices.length > 0 ? 'run' : 'paragraph',
          paragraphIndex: paragraph.index,
          runIndices,
          startCharOffset: startOffset,
          endCharOffset: endOffset,
        };

        matches.push(location);
      }
    }

    if (matches.length === 0) {
      return { found: false, confidence: 0, error: 'No match in paragraphs' };
    }

    if (matches.length === 1) {
      return {
        found: true,
        location: matches[0],
        matchedText: originalAnchor,
        confidence: 1.0,
      };
    }

    // Multiple matches - this is ambiguous
    return {
      found: false,
      confidence: 0,
      error: `Ambiguous: found ${matches.length} matches`,
      alternativeLocations: matches,
    };
  }

  /**
   * Find anchor text in tables
   */
  private findInTables(
    tables: DocxTable[],
    normalizedAnchor: string,
    originalAnchor: string
  ): LocatorResult {
    const matches: ElementLocation[] = [];

    for (const table of tables) {
      for (const row of table.rows) {
        for (const cell of row.cells) {
          for (const paragraph of cell.paragraphs) {
            const normalizedParagraph = paragraph.normalizedText;
            const matchIndex = normalizedParagraph.indexOf(normalizedAnchor);

            if (matchIndex !== -1) {
              const { startOffset, endOffset } = this.mapNormalizedToOriginal(
                paragraph.text,
                normalizedParagraph,
                matchIndex,
                normalizedAnchor.length
              );

              const location: ElementLocation = {
                type: 'table_cell',
                tableIndex: table.index,
                rowIndex: row.index,
                cellIndex: cell.cellIndex,
                paragraphIndex: paragraph.index,
                startCharOffset: startOffset,
                endCharOffset: endOffset,
              };

              matches.push(location);
            }
          }
        }
      }
    }

    if (matches.length === 0) {
      return { found: false, confidence: 0, error: 'No match in tables' };
    }

    if (matches.length === 1) {
      return {
        found: true,
        location: matches[0],
        matchedText: originalAnchor,
        confidence: 1.0,
      };
    }

    return {
      found: false,
      confidence: 0,
      error: `Ambiguous: found ${matches.length} matches in tables`,
      alternativeLocations: matches,
    };
  }

  /**
   * Find anchor text in headers and footers
   */
  private findInHeadersFooters(
    headersFooters: DocxHeaderFooter[],
    normalizedAnchor: string,
    originalAnchor: string
  ): LocatorResult {
    const matches: ElementLocation[] = [];

    for (const hf of headersFooters) {
      for (const paragraph of hf.paragraphs) {
        const normalizedParagraph = paragraph.normalizedText;
        const matchIndex = normalizedParagraph.indexOf(normalizedAnchor);

        if (matchIndex !== -1) {
          const { startOffset, endOffset } = this.mapNormalizedToOriginal(
            paragraph.text,
            normalizedParagraph,
            matchIndex,
            normalizedAnchor.length
          );

          const location: ElementLocation = {
            type: hf.type,
            headerFooterPath: hf.xmlPath,
            paragraphIndex: paragraph.index,
            startCharOffset: startOffset,
            endCharOffset: endOffset,
          };

          matches.push(location);
        }
      }
    }

    if (matches.length === 0) {
      return { found: false, confidence: 0, error: 'No match in headers/footers' };
    }

    if (matches.length === 1) {
      return {
        found: true,
        location: matches[0],
        matchedText: originalAnchor,
        confidence: 1.0,
      };
    }

    return {
      found: false,
      confidence: 0,
      error: `Ambiguous: found ${matches.length} matches in headers/footers`,
      alternativeLocations: matches,
    };
  }

  /**
   * Map normalized text position back to original text position
   */
  private mapNormalizedToOriginal(
    original: string,
    normalized: string,
    normalizedStart: number,
    normalizedLength: number
  ): { startOffset: number; endOffset: number } {
    // Build a mapping from normalized positions to original positions
    let origIndex = 0;
    let normIndex = 0;
    const mapping: number[] = [];

    // Skip leading whitespace in original
    while (origIndex < original.length && /\s/.test(original[origIndex])) {
      origIndex++;
    }

    while (normIndex < normalized.length && origIndex < original.length) {
      const origChar = original[origIndex];
      const normChar = normalized[normIndex];

      if (origChar.toLowerCase() === normChar) {
        mapping[normIndex] = origIndex;
        normIndex++;
        origIndex++;
      } else if (/\s/.test(origChar)) {
        // Original has whitespace, skip it
        origIndex++;
      } else if (normChar === ' ' && /\s/.test(origChar)) {
        // Normalized space matches original whitespace
        mapping[normIndex] = origIndex;
        normIndex++;
        origIndex++;
      } else {
        // Mismatch - advance both
        origIndex++;
      }
    }

    // Fill in any remaining positions
    while (normIndex < normalized.length) {
      mapping[normIndex] = original.length;
      normIndex++;
    }

    const startOffset = mapping[normalizedStart] ?? 0;
    const endNormIndex = normalizedStart + normalizedLength - 1;
    const endOffset = (mapping[endNormIndex] ?? original.length - 1) + 1;

    return { startOffset, endOffset };
  }

  /**
   * Find which runs contain a character range
   */
  private findRunsContainingRange(
    runs: { index: number; startOffset: number; endOffset: number }[],
    startOffset: number,
    endOffset: number
  ): number[] {
    const runIndices: number[] = [];

    for (const run of runs) {
      // Check if this run overlaps with the target range
      if (run.startOffset < endOffset && run.endOffset > startOffset) {
        runIndices.push(run.index);
      }
    }

    return runIndices;
  }

  /**
   * Get the logs from the last locate operation
   */
  getLogs(): string[] {
    return this.logs;
  }
}

/**
 * Batch locate multiple edit specs
 */
export async function batchLocate(
  parsedDocx: ParsedDocx,
  editSpecs: EditSpec[]
): Promise<Map<string, LocatorResult>> {
  const locator = new DocxLocatorAgent();
  const results = new Map<string, LocatorResult>();

  for (const editSpec of editSpecs) {
    const result = await locator.locate(parsedDocx, editSpec);
    results.set(editSpec.changeId, result);
  }

  return results;
}
