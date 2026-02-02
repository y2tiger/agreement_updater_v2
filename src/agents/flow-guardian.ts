/**
 * Flow Guardian (Verification Gate)
 *
 * Re-parses the output DOCX and enforces:
 * - Requested changes are actually present
 * - Targeted text is replaced/inserted as specified
 * - No changes outside allowed scope
 * - No "[object Object]" contamination
 * - DOCX opens cleanly (structure integrity)
 *
 * If ANY check fails → output is BLOCKED
 */

import {
  EditSpec,
  EditResult,
  VerificationResult,
  VerificationCheck,
  VerificationStatus,
  BlockedError,
} from '../types';
import { parseDocx, ParsedDocx, normalizeText, extractPlainText } from '../utils/docx-parser';
import { checkForObjectStringification, verifyDocumentIntegrity } from '../utils/docx-modifier';

interface GuardianContext {
  originalDocx: ParsedDocx;
  modifiedBuffer: Buffer;
  editSpecs: EditSpec[];
  editResults: EditResult[];
}

/**
 * Flow Guardian Agent
 */
export class FlowGuardianAgent {
  private logs: string[] = [];

  private log(message: string): void {
    const timestamp = new Date().toISOString();
    this.logs.push(`[${timestamp}] ${message}`);
  }

  /**
   * Verify the modified document
   */
  async verify(
    originalDocx: ParsedDocx,
    modifiedBuffer: Buffer,
    editSpecs: EditSpec[],
    editResults: EditResult[]
  ): Promise<VerificationResult> {
    this.logs = [];
    this.log('Starting Flow Guardian verification');

    const checks: VerificationCheck[] = [];

    // STEP 1: Parse the modified document
    let modifiedDocx: ParsedDocx;
    try {
      modifiedDocx = await parseDocx(modifiedBuffer);
      checks.push({
        checkName: 'DOCX_PARSE',
        passed: true,
        details: 'Modified document parsed successfully',
      });
      this.log('PASS: Document parses correctly');
    } catch (error) {
      checks.push({
        checkName: 'DOCX_PARSE',
        passed: false,
        details: `Failed to parse modified document: ${error}`,
      });
      this.log('FAIL: Document parsing failed');
      return {
        status: 'FAIL',
        checks,
        overallError: 'Document cannot be parsed after modifications',
      };
    }

    // STEP 2: Check for [object Object] contamination
    const modifiedText = extractPlainText(modifiedDocx);
    if (modifiedText.includes('[object Object]')) {
      checks.push({
        checkName: 'OBJECT_STRINGIFICATION',
        passed: false,
        details: 'Found "[object Object]" in document text',
        evidence: this.extractContext(modifiedText, '[object Object]'),
      });
      this.log('FAIL: [object Object] contamination detected');
      return {
        status: 'FAIL',
        checks,
        overallError: '[object Object] contamination detected',
      };
    }
    checks.push({
      checkName: 'OBJECT_STRINGIFICATION',
      passed: true,
      details: 'No [object Object] contamination found',
    });
    this.log('PASS: No object stringification');

    // STEP 3: Verify document integrity
    const integrity = verifyDocumentIntegrity(modifiedDocx.documentXml);
    if (!integrity.valid) {
      checks.push({
        checkName: 'DOCUMENT_INTEGRITY',
        passed: false,
        details: `Document integrity check failed: ${integrity.error}`,
      });
      this.log(`FAIL: Document integrity: ${integrity.error}`);
      return {
        status: 'FAIL',
        checks,
        overallError: integrity.error,
      };
    }
    checks.push({
      checkName: 'DOCUMENT_INTEGRITY',
      passed: true,
      details: 'Document structure is valid',
    });
    this.log('PASS: Document integrity');

    // STEP 4: Verify each requested change is present
    const successfulEdits = editResults.filter(r => r.success);
    for (const editResult of successfulEdits) {
      const editSpec = editSpecs.find(s => s.changeId === editResult.changeId);
      if (!editSpec) continue;

      const changeCheck = this.verifyChangePresent(
        modifiedDocx,
        editSpec,
        editResult
      );
      checks.push(changeCheck);

      if (!changeCheck.passed) {
        this.log(`FAIL: Change ${editSpec.changeId} not verified`);
      } else {
        this.log(`PASS: Change ${editSpec.changeId} verified`);
      }
    }

    // STEP 5: Check for unintended changes
    const scopeCheck = this.verifyScopeNotExceeded(
      originalDocx,
      modifiedDocx,
      editSpecs
    );
    checks.push(scopeCheck);

    if (!scopeCheck.passed) {
      this.log('FAIL: Scope exceeded - unintended changes detected');
    } else {
      this.log('PASS: No unintended changes');
    }

    // Determine overall result
    const allPassed = checks.every(c => c.passed);
    const status: VerificationStatus = allPassed ? 'PASS' : 'FAIL';

    this.log(`Verification complete: ${status}`);

    return {
      status,
      checks,
      overallError: allPassed ? undefined : 'One or more verification checks failed',
    };
  }

  /**
   * Verify that a specific change is present in the modified document
   */
  private verifyChangePresent(
    modifiedDocx: ParsedDocx,
    editSpec: EditSpec,
    editResult: EditResult
  ): VerificationCheck {
    const modifiedText = extractPlainText(modifiedDocx);
    const normalizedModified = normalizeText(modifiedText);

    // Check if the afterText is present
    const afterTextNormalized = normalizeText(editSpec.afterText);

    // For delete operations, verify the text is gone
    if (editSpec.editType === 'delete') {
      const beforeTextNormalized = normalizeText(editSpec.beforeSnippet);
      const stillPresent = normalizedModified.includes(beforeTextNormalized);

      return {
        checkName: `CHANGE_${editSpec.changeId}`,
        passed: !stillPresent,
        details: stillPresent
          ? 'Deleted text is still present in document'
          : 'Text was successfully deleted',
        evidence: stillPresent
          ? this.extractContext(modifiedText, editSpec.beforeSnippet)
          : undefined,
      };
    }

    // For replace/insert operations, check if afterText is present
    if (afterTextNormalized.length > 0) {
      // Check for key tokens in the afterText (first and last 20 chars)
      const keyTokens = this.extractKeyTokens(afterTextNormalized);
      const allTokensFound = keyTokens.every(token =>
        normalizedModified.includes(token)
      );

      if (!allTokensFound) {
        return {
          checkName: `CHANGE_${editSpec.changeId}`,
          passed: false,
          details: 'Expected text not found in modified document',
          evidence: `Expected: "${editSpec.afterText.substring(0, 100)}..."`,
        };
      }
    }

    // Check that the original text is no longer present (for replace)
    if (editSpec.editType === 'replace') {
      const beforeTextNormalized = normalizeText(editSpec.beforeSnippet);

      // Only check if the before and after are different
      if (beforeTextNormalized !== afterTextNormalized) {
        const stillPresent = normalizedModified.includes(beforeTextNormalized);

        if (stillPresent) {
          // Could be a partial match or the text appears elsewhere
          // Do a more careful check
          const beforeCount = this.countOccurrences(normalizedModified, beforeTextNormalized);
          const originalText = extractPlainText(modifiedDocx);
          const originalNormalized = normalizeText(originalText);
          const originalCount = this.countOccurrences(originalNormalized, beforeTextNormalized);

          // If the count is the same, the replacement might not have happened
          if (beforeCount >= originalCount) {
            return {
              checkName: `CHANGE_${editSpec.changeId}`,
              passed: false,
              details: 'Original text still present - replacement may have failed',
              evidence: this.extractContext(modifiedText, editSpec.beforeSnippet),
            };
          }
        }
      }
    }

    return {
      checkName: `CHANGE_${editSpec.changeId}`,
      passed: true,
      details: 'Change successfully applied and verified',
    };
  }

  /**
   * Verify no unintended changes were made
   */
  private verifyScopeNotExceeded(
    originalDocx: ParsedDocx,
    modifiedDocx: ParsedDocx,
    editSpecs: EditSpec[]
  ): VerificationCheck {
    // Get the set of paragraphs that should be modified
    const expectedChangedParagraphs = new Set<number>();
    // This is a simplified check - in production would need more sophisticated tracking

    // Compare paragraph counts
    const originalCount = originalDocx.paragraphs.length;
    const modifiedCount = modifiedDocx.paragraphs.length;

    // Allow for inserted/deleted paragraphs based on edit specs
    const insertCount = editSpecs.filter(
      s => s.editType === 'insert_before' || s.editType === 'insert_after'
    ).length;
    const deleteCount = editSpecs.filter(s => s.editType === 'delete').length;

    const expectedDiff = insertCount - deleteCount;
    const actualDiff = modifiedCount - originalCount;

    if (actualDiff !== expectedDiff) {
      return {
        checkName: 'SCOPE_CHECK',
        passed: false,
        details: `Unexpected paragraph count change. Expected diff: ${expectedDiff}, Actual: ${actualDiff}`,
      };
    }

    // Compare paragraphs that shouldn't have changed
    let unexpectedChanges = 0;
    const minCount = Math.min(originalCount, modifiedCount);

    // This is a simplified comparison - in production would track exact positions
    for (let i = 0; i < minCount && i < 100; i++) {
      const originalPara = originalDocx.paragraphs[i];
      const modifiedPara = modifiedDocx.paragraphs[i];

      if (!originalPara || !modifiedPara) continue;

      // Check if this paragraph was supposed to be modified
      const wasTargeted = editSpecs.some(spec => {
        // Simplified check - would need location mapping in production
        return normalizeText(originalPara.text).includes(
          normalizeText(spec.anchorText).substring(0, 50)
        );
      });

      if (!wasTargeted && originalPara.normalizedText !== modifiedPara.normalizedText) {
        unexpectedChanges++;
      }
    }

    if (unexpectedChanges > 0) {
      return {
        checkName: 'SCOPE_CHECK',
        passed: false,
        details: `Found ${unexpectedChanges} paragraphs changed that were not targeted`,
      };
    }

    return {
      checkName: 'SCOPE_CHECK',
      passed: true,
      details: 'No unintended changes detected',
    };
  }

  /**
   * Extract context around a match
   */
  private extractContext(text: string, search: string, contextLength: number = 50): string {
    const index = text.indexOf(search);
    if (index === -1) return '';

    const start = Math.max(0, index - contextLength);
    const end = Math.min(text.length, index + search.length + contextLength);

    return `...${text.substring(start, end)}...`;
  }

  /**
   * Extract key tokens from text for verification
   */
  private extractKeyTokens(text: string): string[] {
    const tokens: string[] = [];

    // Get first 30 chars (if long enough)
    if (text.length >= 30) {
      tokens.push(text.substring(0, 30));
    } else if (text.length >= 10) {
      tokens.push(text.substring(0, 10));
    }

    // Get last 30 chars (if long enough and different from first)
    if (text.length >= 60) {
      tokens.push(text.substring(text.length - 30));
    }

    // If text is short, use the whole thing
    if (tokens.length === 0 && text.length > 0) {
      tokens.push(text);
    }

    return tokens;
  }

  /**
   * Count occurrences of a substring
   */
  private countOccurrences(text: string, search: string): number {
    let count = 0;
    let pos = 0;

    while ((pos = text.indexOf(search, pos)) !== -1) {
      count++;
      pos += search.length;
    }

    return count;
  }

  /**
   * Get the logs from the last verification
   */
  getLogs(): string[] {
    return this.logs;
  }
}
