/**
 * QA Evidence Agent
 *
 * Outputs BEFORE/AFTER snippets per change.
 * Classifies failures:
 * - Target not found
 * - Ambiguous match
 * - Apply failed
 * - Verification failed
 * - DOCX integrity failed
 */

import {
  ChangeItem,
  EditSpec,
  EditResult,
  QAEvidence,
  VerificationResult,
  VerificationCheck,
  FailureClassification,
  LocatorResult,
} from '../types';

interface EvidenceContext {
  changeItems: ChangeItem[];
  editResults: EditResult[];
  locatorResults: Map<string, LocatorResult>;
  verificationResult: VerificationResult;
}

/**
 * QA Evidence Agent
 */
export class QAEvidenceAgent {
  private logs: string[] = [];

  private log(message: string): void {
    const timestamp = new Date().toISOString();
    this.logs.push(`[${timestamp}] ${message}`);
  }

  /**
   * Generate QA evidence for all changes
   */
  generateEvidence(
    changeItems: ChangeItem[],
    editResults: EditResult[],
    locatorResults: Map<string, LocatorResult>,
    verificationResult: VerificationResult
  ): QAEvidence[] {
    this.logs = [];
    this.log('Generating QA evidence');

    const evidence: QAEvidence[] = [];

    for (const changeItem of changeItems) {
      const ev = this.generateEvidenceForChange(
        changeItem,
        editResults,
        locatorResults,
        verificationResult
      );
      evidence.push(ev);
      this.log(`Evidence for ${changeItem.changeId}: ${ev.verificationStatus}`);
    }

    this.log(`Generated ${evidence.length} evidence records`);
    return evidence;
  }

  /**
   * Generate evidence for a single change
   */
  private generateEvidenceForChange(
    changeItem: ChangeItem,
    editResults: EditResult[],
    locatorResults: Map<string, LocatorResult>,
    verificationResult: VerificationResult
  ): QAEvidence {
    const changeId = changeItem.changeId;
    const editSpec = changeItem.editSpec;
    const locatorResult = locatorResults.get(changeId);
    const editResult = editResults.find(r => r.changeId === changeId);

    // Determine the before snippet
    let beforeSnippet = '';
    if (editSpec) {
      beforeSnippet = editSpec.beforeSnippet;
    } else if (changeItem.selectedTarget) {
      beforeSnippet = changeItem.selectedTarget.anchorText;
    }

    // Determine the after snippet
    let afterSnippet = '';
    if (editResult?.success && editSpec) {
      afterSnippet = editSpec.afterText;
    }

    // Classify the result
    const { status, classification, details } = this.classifyResult(
      changeItem,
      locatorResult,
      editResult,
      verificationResult
    );

    return {
      changeId,
      beforeSnippet: this.truncateSnippet(beforeSnippet),
      afterSnippet: this.truncateSnippet(afterSnippet),
      verificationStatus: status,
      failureClassification: classification,
      failureDetails: details,
    };
  }

  /**
   * Classify the result of a change
   */
  private classifyResult(
    changeItem: ChangeItem,
    locatorResult: LocatorResult | undefined,
    editResult: EditResult | undefined,
    verificationResult: VerificationResult
  ): {
    status: 'PASS' | 'FAIL';
    classification?: FailureClassification;
    details?: string;
  } {
    // Check if change needed confirmation
    if (changeItem.status === 'NEEDS_CONFIRMATION') {
      if (changeItem.failureReason?.includes('Ambiguous')) {
        return {
          status: 'FAIL',
          classification: 'AMBIGUOUS_MATCH',
          details: changeItem.failureReason,
        };
      }
      return {
        status: 'FAIL',
        classification: 'TARGET_NOT_FOUND',
        details: changeItem.failureReason || 'Change requires confirmation',
      };
    }

    // Check locator result
    if (!locatorResult || !locatorResult.found) {
      if (locatorResult?.alternativeLocations && locatorResult.alternativeLocations.length > 1) {
        return {
          status: 'FAIL',
          classification: 'AMBIGUOUS_MATCH',
          details: locatorResult.error || 'Multiple matches found',
        };
      }
      return {
        status: 'FAIL',
        classification: 'TARGET_NOT_FOUND',
        details: locatorResult?.error || 'Target location not found',
      };
    }

    // Check edit result
    if (!editResult) {
      return {
        status: 'FAIL',
        classification: 'APPLY_FAILED',
        details: 'Edit was not executed',
      };
    }

    if (!editResult.success) {
      if (editResult.error?.includes('[object Object]')) {
        return {
          status: 'FAIL',
          classification: 'OBJECT_STRINGIFICATION',
          details: editResult.error,
        };
      }
      if (editResult.error?.includes('Integrity')) {
        return {
          status: 'FAIL',
          classification: 'DOCX_INTEGRITY_FAILED',
          details: editResult.error,
        };
      }
      return {
        status: 'FAIL',
        classification: 'APPLY_FAILED',
        details: editResult.error,
      };
    }

    // Check verification result for this change
    const changeVerification = verificationResult.checks.find(
      c => c.checkName === `CHANGE_${changeItem.changeId}`
    );

    if (changeVerification && !changeVerification.passed) {
      return {
        status: 'FAIL',
        classification: 'VERIFICATION_FAILED',
        details: changeVerification.details,
      };
    }

    // Check overall verification
    if (verificationResult.status === 'FAIL') {
      // Check if it's an integrity failure
      const integrityCheck = verificationResult.checks.find(
        c => c.checkName === 'DOCUMENT_INTEGRITY' && !c.passed
      );
      if (integrityCheck) {
        return {
          status: 'FAIL',
          classification: 'DOCX_INTEGRITY_FAILED',
          details: integrityCheck.details,
        };
      }

      // Check for object stringification
      const objectCheck = verificationResult.checks.find(
        c => c.checkName === 'OBJECT_STRINGIFICATION' && !c.passed
      );
      if (objectCheck) {
        return {
          status: 'FAIL',
          classification: 'OBJECT_STRINGIFICATION',
          details: objectCheck.details,
        };
      }

      return {
        status: 'FAIL',
        classification: 'VERIFICATION_FAILED',
        details: verificationResult.overallError,
      };
    }

    // All checks passed
    return {
      status: 'PASS',
    };
  }

  /**
   * Truncate a snippet for display
   */
  private truncateSnippet(text: string, maxLength: number = 200): string {
    if (text.length <= maxLength) {
      return text;
    }
    return text.substring(0, maxLength - 3) + '...';
  }

  /**
   * Generate a summary report
   */
  generateSummaryReport(evidence: QAEvidence[]): string {
    const passed = evidence.filter(e => e.verificationStatus === 'PASS').length;
    const failed = evidence.filter(e => e.verificationStatus === 'FAIL').length;

    let report = `QA Evidence Summary\n`;
    report += `==================\n\n`;
    report += `Total Changes: ${evidence.length}\n`;
    report += `Passed: ${passed}\n`;
    report += `Failed: ${failed}\n\n`;

    if (failed > 0) {
      report += `Failure Details:\n`;
      report += `----------------\n`;

      const byClassification = new Map<string, QAEvidence[]>();
      for (const ev of evidence) {
        if (ev.failureClassification) {
          const list = byClassification.get(ev.failureClassification) || [];
          list.push(ev);
          byClassification.set(ev.failureClassification, list);
        }
      }

      for (const [classification, items] of byClassification) {
        report += `\n${classification} (${items.length}):\n`;
        for (const item of items) {
          report += `  - ${item.changeId}: ${item.failureDetails || 'No details'}\n`;
        }
      }
    }

    report += `\nChange Details:\n`;
    report += `--------------\n`;

    for (const ev of evidence) {
      report += `\n[${ev.changeId}] ${ev.verificationStatus}\n`;
      if (ev.beforeSnippet) {
        report += `  BEFORE: "${ev.beforeSnippet}"\n`;
      }
      if (ev.afterSnippet) {
        report += `  AFTER:  "${ev.afterSnippet}"\n`;
      }
      if (ev.failureDetails) {
        report += `  ERROR:  ${ev.failureDetails}\n`;
      }
    }

    return report;
  }

  /**
   * Get the logs from the last operation
   */
  getLogs(): string[] {
    return this.logs;
  }
}
