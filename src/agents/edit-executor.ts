/**
 * Edit Executor Agent
 *
 * Applies changes ONLY within Edit Spec boundaries.
 * Preserves numbering, styles, tables, headers/footers.
 * Applies yellow highlight to modified runs only.
 * Guarantees no "[object Object]" or object stringification.
 */

import {
  EditSpec,
  ElementLocation,
  EditResult,
  LocatorResult,
  ContractEditError,
  BlockedError,
} from '../types';
import { ParsedDocx, saveDocx, serializeXml } from '../utils/docx-parser';
import {
  applyEdit,
  checkForObjectStringification,
  verifyDocumentIntegrity,
} from '../utils/docx-modifier';

interface ExecutionContext {
  parsedDocx: ParsedDocx;
  modifiedXml: unknown;
  appliedEdits: EditResult[];
  logs: string[];
}

/**
 * Edit Executor Agent
 */
export class EditExecutorAgent {
  private logs: string[] = [];

  private log(message: string): void {
    const timestamp = new Date().toISOString();
    this.logs.push(`[${timestamp}] ${message}`);
  }

  /**
   * Execute a batch of edits
   */
  async execute(
    parsedDocx: ParsedDocx,
    editSpecs: EditSpec[],
    locations: Map<string, LocatorResult>
  ): Promise<{
    success: boolean;
    modifiedBuffer?: Buffer;
    results: EditResult[];
    logs: string[];
  }> {
    this.logs = [];
    this.log(`Starting execution of ${editSpecs.length} edits`);

    // Start with the original document XML
    let currentXml = JSON.parse(JSON.stringify(parsedDocx.documentXml));
    const results: EditResult[] = [];
    let hasFailure = false;

    // Apply edits sequentially to maintain consistency
    for (const editSpec of editSpecs) {
      const locatorResult = locations.get(editSpec.changeId);

      if (!locatorResult || !locatorResult.found || !locatorResult.location) {
        this.log(`Skipping change ${editSpec.changeId}: location not found`);
        results.push({
          changeId: editSpec.changeId,
          success: false,
          beforeText: editSpec.beforeSnippet,
          afterText: '',
          location: {} as ElementLocation,
          error: 'Location not found',
        });
        hasFailure = true;
        continue;
      }

      this.log(`Applying change ${editSpec.changeId}`);
      this.log(`  Edit type: ${editSpec.editType}`);
      this.log(`  Target: ${locatorResult.location.type} at index ${locatorResult.location.paragraphIndex}`);

      try {
        // Create a modified copy of parsed docx with current XML
        const workingDocx: ParsedDocx = {
          ...parsedDocx,
          documentXml: currentXml,
        };

        const editResult = await applyEdit(workingDocx, editSpec, locatorResult.location);

        if (!editResult.success) {
          this.log(`  Edit failed: ${editResult.error}`);
          results.push({
            changeId: editSpec.changeId,
            success: false,
            beforeText: editSpec.beforeSnippet,
            afterText: '',
            location: locatorResult.location,
            error: editResult.error,
          });
          hasFailure = true;
          continue;
        }

        // Check for object stringification immediately
        if (checkForObjectStringification(editResult.modifiedXml)) {
          this.log(`  Edit failed: [object Object] contamination detected`);
          results.push({
            changeId: editSpec.changeId,
            success: false,
            beforeText: editSpec.beforeSnippet,
            afterText: '',
            location: locatorResult.location,
            error: '[object Object] contamination detected',
          });
          hasFailure = true;
          continue;
        }

        // Verify document integrity after this edit
        const integrity = verifyDocumentIntegrity(editResult.modifiedXml);
        if (!integrity.valid) {
          this.log(`  Edit failed integrity check: ${integrity.error}`);
          results.push({
            changeId: editSpec.changeId,
            success: false,
            beforeText: editSpec.beforeSnippet,
            afterText: '',
            location: locatorResult.location,
            error: `Integrity check failed: ${integrity.error}`,
          });
          hasFailure = true;
          continue;
        }

        // Update the current XML for the next edit
        currentXml = editResult.modifiedXml;

        this.log(`  Edit applied successfully`);
        results.push({
          changeId: editSpec.changeId,
          success: true,
          beforeText: editSpec.beforeSnippet,
          afterText: editSpec.afterText,
          location: locatorResult.location,
        });
      } catch (error) {
        this.log(`  Edit threw error: ${error}`);
        results.push({
          changeId: editSpec.changeId,
          success: false,
          beforeText: editSpec.beforeSnippet,
          afterText: '',
          location: locatorResult.location,
          error: error instanceof Error ? error.message : 'Unknown error',
        });
        hasFailure = true;
      }
    }

    // Final integrity check on the entire document
    const finalIntegrity = verifyDocumentIntegrity(currentXml);
    if (!finalIntegrity.valid) {
      this.log(`Final document failed integrity check: ${finalIntegrity.error}`);
      return {
        success: false,
        results,
        logs: this.logs,
      };
    }

    // If any edit failed, we don't produce an output
    if (hasFailure) {
      this.log('Some edits failed - not producing output');
      return {
        success: false,
        results,
        logs: this.logs,
      };
    }

    // Save the modified document
    try {
      const modifiedBuffer = saveDocx(parsedDocx, currentXml);
      this.log('Document saved successfully');

      return {
        success: true,
        modifiedBuffer,
        results,
        logs: this.logs,
      };
    } catch (error) {
      this.log(`Failed to save document: ${error}`);
      return {
        success: false,
        results,
        logs: this.logs,
      };
    }
  }

  /**
   * Execute a single edit (for retry scenarios)
   */
  async executeSingle(
    parsedDocx: ParsedDocx,
    editSpec: EditSpec,
    location: ElementLocation
  ): Promise<EditResult & { modifiedXml?: unknown }> {
    this.log(`Executing single edit ${editSpec.changeId}`);

    try {
      const editResult = await applyEdit(parsedDocx, editSpec, location);

      if (!editResult.success) {
        return {
          changeId: editSpec.changeId,
          success: false,
          beforeText: editSpec.beforeSnippet,
          afterText: '',
          location,
          error: editResult.error,
        };
      }

      if (checkForObjectStringification(editResult.modifiedXml)) {
        return {
          changeId: editSpec.changeId,
          success: false,
          beforeText: editSpec.beforeSnippet,
          afterText: '',
          location,
          error: '[object Object] contamination detected',
        };
      }

      const integrity = verifyDocumentIntegrity(editResult.modifiedXml);
      if (!integrity.valid) {
        return {
          changeId: editSpec.changeId,
          success: false,
          beforeText: editSpec.beforeSnippet,
          afterText: '',
          location,
          error: `Integrity check failed: ${integrity.error}`,
        };
      }

      return {
        changeId: editSpec.changeId,
        success: true,
        beforeText: editSpec.beforeSnippet,
        afterText: editSpec.afterText,
        location,
        modifiedXml: editResult.modifiedXml,
      };
    } catch (error) {
      return {
        changeId: editSpec.changeId,
        success: false,
        beforeText: editSpec.beforeSnippet,
        afterText: '',
        location,
        error: error instanceof Error ? error.message : 'Unknown error',
      };
    }
  }

  /**
   * Get the logs from the last execution
   */
  getLogs(): string[] {
    return this.logs;
  }
}
