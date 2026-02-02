/**
 * Main Pipeline Orchestrator
 *
 * Coordinates all agents through the execution pipeline:
 * PHASE 0 — Global Context Scan (OpenAI, read-only)
 * PHASE 1 — Change Decomposition (OpenAI)
 * PHASE 2 — Candidate Targeting (OpenAI)
 * PHASE 3 — Edit Spec Lock (OpenAI)
 * PHASE 4 — Apply (Code)
 * PHASE 5 — Post-Verification (Guardian)
 */

import {
  PipelineRequest,
  PipelineResponse,
  PipelineLog,
  ChangeItem,
  EditSpec,
  QAEvidence,
  VerificationResult,
  LocatorResult,
  BlockedError,
  ContractEditError,
} from '../types';
import { parseDocx, ParsedDocx } from '../utils/docx-parser';
import { ContractArchitectAgent } from '../agents/contract-architect';
import { DocxLocatorAgent } from '../agents/docx-locator';
import { EditExecutorAgent } from '../agents/edit-executor';
import { FlowGuardianAgent } from '../agents/flow-guardian';
import { QAEvidenceAgent } from '../agents/qa-evidence';

const MAX_RETRIES = 2;
const PIPELINE_TIMEOUT = 120000; // 2 minutes

interface PipelineConfig {
  openaiApiKey: string;
  openaiModel?: string;
  timeout?: number;
}

/**
 * Contract Edit Pipeline
 */
export class ContractEditPipeline {
  private config: PipelineConfig;
  private logs: PipelineLog[] = [];

  constructor(config: PipelineConfig) {
    this.config = config;
  }

  private log(phase: string, message: string, data?: unknown): void {
    this.logs.push({
      phase,
      timestamp: new Date(),
      message,
      data,
    });
  }

  /**
   * Execute the full pipeline
   */
  async execute(request: PipelineRequest): Promise<PipelineResponse> {
    this.logs = [];
    const startTime = Date.now();
    const timeout = this.config.timeout || PIPELINE_TIMEOUT;

    this.log('INIT', `Starting pipeline for ${request.fileName}`);
    this.log('INIT', `User instruction: ${request.userInstruction}`);

    try {
      // Check timeout periodically
      const checkTimeout = () => {
        if (Date.now() - startTime > timeout) {
          throw new BlockedError('Pipeline timeout exceeded', 'TIMEOUT');
        }
      };

      // STEP 1: Parse the input document
      this.log('PARSE', 'Parsing input DOCX');
      const parsedDocx = await parseDocx(request.documentBuffer);
      this.log('PARSE', `Parsed ${parsedDocx.paragraphs.length} paragraphs, ${parsedDocx.tables.length} tables`);

      checkTimeout();

      // STEP 2: Run Contract Architect (Phases 0-3)
      this.log('ARCHITECT', 'Running Contract Architect');
      const architect = new ContractArchitectAgent(
        this.config.openaiApiKey,
        this.config.openaiModel
      );

      const architectResult = await architect.analyze(parsedDocx, request.userInstruction);

      this.log('ARCHITECT', `Document type: ${architectResult.documentContext.documentType}`);
      this.log('ARCHITECT', `Change items: ${architectResult.changeItems.length}`);

      for (const log of architectResult.logs) {
        this.log('ARCHITECT', log);
      }

      checkTimeout();

      // Filter to only high-confidence changes
      const executableChanges = architectResult.changeItems.filter(
        c => c.status === 'HIGH_CONFIDENCE' && c.editSpec
      );

      this.log('ARCHITECT', `Executable changes: ${executableChanges.length}`);

      if (executableChanges.length === 0) {
        // No executable changes - return early
        const qaAgent = new QAEvidenceAgent();
        const qaEvidence = qaAgent.generateEvidence(
          architectResult.changeItems,
          [],
          new Map(),
          { status: 'FAIL', checks: [] }
        );

        return {
          success: false,
          changeItems: architectResult.changeItems,
          qaEvidence,
          verificationResult: {
            status: 'FAIL',
            checks: [],
            overallError: 'No changes could be executed with high confidence',
          },
          blockedReason: 'No changes could be executed. Changes may need confirmation.',
          logs: this.logs,
        };
      }

      // STEP 3: Run DOCX Locator for each executable change
      this.log('LOCATOR', 'Running DOCX Locator');
      const locator = new DocxLocatorAgent();
      const locatorResults = new Map<string, LocatorResult>();

      for (const change of executableChanges) {
        if (!change.editSpec) continue;

        const result = await locator.locate(parsedDocx, change.editSpec);
        locatorResults.set(change.changeId, result);

        this.log('LOCATOR', `Change ${change.changeId}: ${result.found ? 'FOUND' : 'NOT FOUND'}`);

        if (!result.found) {
          change.status = 'FAILED';
          change.failureReason = result.error;
        }
      }

      for (const log of locator.getLogs()) {
        this.log('LOCATOR', log);
      }

      checkTimeout();

      // STEP 4: Execute edits
      this.log('EXECUTOR', 'Running Edit Executor');
      const executor = new EditExecutorAgent();

      // Get edit specs for located changes
      const editSpecs = executableChanges
        .filter(c => c.editSpec && locatorResults.get(c.changeId)?.found)
        .map(c => c.editSpec!);

      if (editSpecs.length === 0) {
        const qaAgent = new QAEvidenceAgent();
        const qaEvidence = qaAgent.generateEvidence(
          architectResult.changeItems,
          [],
          locatorResults,
          { status: 'FAIL', checks: [] }
        );

        return {
          success: false,
          changeItems: architectResult.changeItems,
          qaEvidence,
          verificationResult: {
            status: 'FAIL',
            checks: [],
            overallError: 'No changes could be located in the document',
          },
          blockedReason: 'Could not locate target text in document',
          logs: this.logs,
        };
      }

      const executionResult = await executor.execute(parsedDocx, editSpecs, locatorResults);

      for (const log of executor.getLogs()) {
        this.log('EXECUTOR', log);
      }

      // Update change item statuses based on execution results
      for (const result of executionResult.results) {
        const change = architectResult.changeItems.find(c => c.changeId === result.changeId);
        if (change) {
          if (result.success) {
            change.status = 'APPLIED';
          } else {
            change.status = 'FAILED';
            change.failureReason = result.error;
          }
        }
      }

      checkTimeout();

      if (!executionResult.success || !executionResult.modifiedBuffer) {
        // Execution failed
        this.log('EXECUTOR', 'Execution failed - attempting retry');

        // Attempt retry with alternative targets
        const retryResult = await this.retryWithAlternatives(
          parsedDocx,
          architectResult.changeItems,
          locatorResults,
          executionResult.results
        );

        if (retryResult) {
          // Update with retry results
          const { modifiedBuffer, results } = retryResult;
          executionResult.modifiedBuffer = modifiedBuffer;
          executionResult.results = results;
          executionResult.success = true;
        } else {
          const qaAgent = new QAEvidenceAgent();
          const qaEvidence = qaAgent.generateEvidence(
            architectResult.changeItems,
            executionResult.results,
            locatorResults,
            { status: 'FAIL', checks: [] }
          );

          return {
            success: false,
            changeItems: architectResult.changeItems,
            qaEvidence,
            verificationResult: {
              status: 'FAIL',
              checks: [],
              overallError: 'Edit execution failed',
            },
            blockedReason: 'Failed to apply edits to document',
            logs: this.logs,
          };
        }
      }

      // STEP 5: Run Flow Guardian verification
      this.log('GUARDIAN', 'Running Flow Guardian verification');
      const guardian = new FlowGuardianAgent();

      const verificationResult = await guardian.verify(
        parsedDocx,
        executionResult.modifiedBuffer!,
        editSpecs,
        executionResult.results
      );

      for (const log of guardian.getLogs()) {
        this.log('GUARDIAN', log);
      }

      // STEP 6: Generate QA Evidence
      this.log('QA', 'Generating QA evidence');
      const qaAgent = new QAEvidenceAgent();
      const qaEvidence = qaAgent.generateEvidence(
        architectResult.changeItems,
        executionResult.results,
        locatorResults,
        verificationResult
      );

      // Check if verification passed
      if (verificationResult.status !== 'PASS') {
        this.log('GUARDIAN', `Verification FAILED: ${verificationResult.overallError}`);

        return {
          success: false,
          changeItems: architectResult.changeItems,
          qaEvidence,
          verificationResult,
          blockedReason: verificationResult.overallError || 'Verification failed',
          logs: this.logs,
        };
      }

      // SUCCESS
      this.log('COMPLETE', 'Pipeline completed successfully');

      return {
        success: true,
        documentBuffer: executionResult.modifiedBuffer,
        changeItems: architectResult.changeItems,
        qaEvidence,
        verificationResult,
        logs: this.logs,
      };
    } catch (error) {
      this.log('ERROR', `Pipeline error: ${error}`);

      if (error instanceof BlockedError) {
        return {
          success: false,
          changeItems: [],
          qaEvidence: [],
          verificationResult: {
            status: 'FAIL',
            checks: [],
            overallError: error.message,
          },
          blockedReason: error.message,
          logs: this.logs,
        };
      }

      return {
        success: false,
        changeItems: [],
        qaEvidence: [],
        verificationResult: {
          status: 'FAIL',
          checks: [],
          overallError: error instanceof Error ? error.message : 'Unknown error',
        },
        blockedReason: error instanceof Error ? error.message : 'Unknown pipeline error',
        logs: this.logs,
      };
    }
  }

  /**
   * Retry failed edits with alternative targets
   */
  private async retryWithAlternatives(
    parsedDocx: ParsedDocx,
    changeItems: ChangeItem[],
    locatorResults: Map<string, LocatorResult>,
    previousResults: { changeId: string; success: boolean }[]
  ): Promise<{ modifiedBuffer: Buffer; results: any[] } | null> {
    this.log('RETRY', 'Attempting retry with alternative targets');

    const failedChangeIds = previousResults
      .filter(r => !r.success)
      .map(r => r.changeId);

    if (failedChangeIds.length === 0) {
      return null;
    }

    // Try alternative locations
    for (const changeId of failedChangeIds) {
      const locatorResult = locatorResults.get(changeId);
      if (locatorResult?.alternativeLocations && locatorResult.alternativeLocations.length > 0) {
        // Try the first alternative
        const alternative = locatorResult.alternativeLocations[0];
        this.log('RETRY', `Trying alternative location for ${changeId}`);

        // Update the locator result with the alternative
        locatorResults.set(changeId, {
          ...locatorResult,
          found: true,
          location: alternative,
        });
      }
    }

    // Re-run executor
    const executor = new EditExecutorAgent();
    const editSpecs = changeItems
      .filter(c => c.editSpec && locatorResults.get(c.changeId)?.found)
      .map(c => c.editSpec!);

    const retryResult = await executor.execute(parsedDocx, editSpecs, locatorResults);

    if (retryResult.success && retryResult.modifiedBuffer) {
      this.log('RETRY', 'Retry succeeded');
      return {
        modifiedBuffer: retryResult.modifiedBuffer,
        results: retryResult.results,
      };
    }

    this.log('RETRY', 'Retry failed');
    return null;
  }

  /**
   * Get pipeline logs
   */
  getLogs(): PipelineLog[] {
    return this.logs;
  }
}
