/**
 * Edit API Routes
 *
 * Handles document upload and edit processing.
 * All files processed in /tmp directories.
 * Any failure returns a clear error response (never partial success).
 * "Edit completed successfully" returned ONLY after full verification PASS.
 */

import { Router, Request, Response } from 'express';
import multer from 'multer';
import { v4 as uuidv4 } from 'uuid';
import * as fs from 'fs';
import * as path from 'path';
import { ContractEditPipeline } from '../pipeline';
import {
  EditRequest,
  EditResponse,
  ChangeItemSummary,
  PipelineResponse,
} from '../types';

const router = Router();

// Configure multer for file uploads to /tmp
const storage = multer.diskStorage({
  destination: (req, file, cb) => {
    const uploadDir = '/tmp/contract-edits';
    if (!fs.existsSync(uploadDir)) {
      fs.mkdirSync(uploadDir, { recursive: true });
    }
    cb(null, uploadDir);
  },
  filename: (req, file, cb) => {
    const uniqueId = uuidv4();
    const ext = path.extname(file.originalname);
    cb(null, `${uniqueId}${ext}`);
  },
});

const upload = multer({
  storage,
  limits: {
    fileSize: 10 * 1024 * 1024, // 10MB limit
  },
  fileFilter: (req, file, cb) => {
    if (
      file.mimetype ===
        'application/vnd.openxmlformats-officedocument.wordprocessingml.document' ||
      file.originalname.endsWith('.docx')
    ) {
      cb(null, true);
    } else {
      cb(new Error('Only .docx files are allowed'));
    }
  },
});

// Store processed files temporarily for download
const processedFiles = new Map<string, { buffer: Buffer; fileName: string; expiry: number }>();

// Clean up expired files periodically
setInterval(() => {
  const now = Date.now();
  for (const [id, file] of processedFiles) {
    if (file.expiry < now) {
      processedFiles.delete(id);
    }
  }
}, 60000); // Check every minute

/**
 * POST /api/edit
 *
 * Upload a DOCX file and apply edits based on natural language instructions.
 *
 * Body (multipart/form-data):
 * - file: The DOCX file to edit
 * - instruction: Natural language edit instruction
 *
 * Response:
 * - success: boolean
 * - downloadUrl: URL to download the modified file (if success)
 * - changeItems: Summary of changes
 * - qaEvidence: Evidence for each change
 * - verificationResult: Verification status
 * - blockedReason: Reason for failure (if not success)
 */
router.post('/edit', upload.single('file'), async (req: Request, res: Response) => {
  const requestId = uuidv4();
  const startTime = Date.now();

  console.log(`[${requestId}] Edit request received`);

  let uploadedFilePath: string | undefined;

  try {
    // Validate request
    if (!req.file) {
      return res.status(400).json({
        success: false,
        blockedReason: 'No file uploaded',
      });
    }

    const instruction = req.body.instruction as string;
    if (!instruction || instruction.trim().length === 0) {
      // Clean up uploaded file
      fs.unlinkSync(req.file.path);
      return res.status(400).json({
        success: false,
        blockedReason: 'No instruction provided',
      });
    }

    uploadedFilePath = req.file.path;
    const fileName = req.file.originalname;

    console.log(`[${requestId}] File: ${fileName}`);
    console.log(`[${requestId}] Instruction: ${instruction}`);

    // Check for OpenAI API key
    const openaiApiKey = process.env.OPENAI_API_KEY;
    if (!openaiApiKey) {
      fs.unlinkSync(uploadedFilePath);
      return res.status(500).json({
        success: false,
        blockedReason: 'OpenAI API key not configured',
      });
    }

    // Read the uploaded file
    const documentBuffer = fs.readFileSync(uploadedFilePath);

    // Run the pipeline
    const pipeline = new ContractEditPipeline({
      openaiApiKey,
      openaiModel: process.env.OPENAI_MODEL || 'gpt-4-turbo-preview',
      timeout: parseInt(process.env.PIPELINE_TIMEOUT || '120000', 10),
    });

    const result = await pipeline.execute({
      documentBuffer,
      userInstruction: instruction,
      fileName,
    });

    // Clean up uploaded file
    fs.unlinkSync(uploadedFilePath);
    uploadedFilePath = undefined;

    // Build response
    const response: EditResponse = {
      success: result.success,
      changeItems: result.changeItems.map(summarizeChangeItem),
      qaEvidence: result.qaEvidence,
      verificationResult: result.verificationResult,
      blockedReason: result.blockedReason,
    };

    if (result.success && result.documentBuffer) {
      // Store the modified file for download
      const downloadId = uuidv4();
      const modifiedFileName = `edited_${fileName}`;

      processedFiles.set(downloadId, {
        buffer: result.documentBuffer,
        fileName: modifiedFileName,
        expiry: Date.now() + 15 * 60 * 1000, // 15 minutes
      });

      response.downloadUrl = `/api/download/${downloadId}`;

      console.log(`[${requestId}] Edit completed successfully in ${Date.now() - startTime}ms`);
    } else {
      console.log(`[${requestId}] Edit blocked: ${result.blockedReason}`);
    }

    // Log the result
    logPipelineResult(requestId, result);

    return res.json(response);
  } catch (error) {
    console.error(`[${requestId}] Pipeline error:`, error);

    // Clean up uploaded file on error
    if (uploadedFilePath && fs.existsSync(uploadedFilePath)) {
      fs.unlinkSync(uploadedFilePath);
    }

    return res.status(500).json({
      success: false,
      blockedReason: error instanceof Error ? error.message : 'Internal server error',
      changeItems: [],
      qaEvidence: [],
      verificationResult: {
        status: 'FAIL',
        checks: [],
        overallError: 'Pipeline execution failed',
      },
    });
  }
});

/**
 * GET /api/download/:id
 *
 * Download a processed DOCX file.
 */
router.get('/download/:id', (req: Request, res: Response) => {
  const { id } = req.params;

  const file = processedFiles.get(id);
  if (!file) {
    return res.status(404).json({
      success: false,
      error: 'File not found or expired',
    });
  }

  // Check expiry
  if (file.expiry < Date.now()) {
    processedFiles.delete(id);
    return res.status(410).json({
      success: false,
      error: 'Download link has expired',
    });
  }

  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
  res.setHeader('Content-Disposition', `attachment; filename="${file.fileName}"`);
  res.send(file.buffer);

  // Remove after download
  processedFiles.delete(id);
});

/**
 * GET /api/health
 *
 * Health check endpoint.
 */
router.get('/health', (req: Request, res: Response) => {
  res.json({
    status: 'ok',
    timestamp: new Date().toISOString(),
    openaiConfigured: !!process.env.OPENAI_API_KEY,
  });
});

/**
 * Summarize a change item for the response
 */
function summarizeChangeItem(item: any): ChangeItemSummary {
  return {
    changeId: item.changeId,
    description: item.description,
    status: item.status,
    beforeSnippet: item.editSpec?.beforeSnippet?.substring(0, 200),
    afterSnippet: item.editSpec?.afterText?.substring(0, 200),
  };
}

/**
 * Log pipeline result for debugging
 */
function logPipelineResult(requestId: string, result: PipelineResponse): void {
  console.log(`[${requestId}] Pipeline Result:`);
  console.log(`  Success: ${result.success}`);
  console.log(`  Changes: ${result.changeItems.length}`);

  if (result.blockedReason) {
    console.log(`  Blocked: ${result.blockedReason}`);
  }

  console.log(`  Verification: ${result.verificationResult.status}`);

  for (const check of result.verificationResult.checks) {
    console.log(`    ${check.checkName}: ${check.passed ? 'PASS' : 'FAIL'}`);
    if (!check.passed) {
      console.log(`      ${check.details}`);
    }
  }

  // Log QA evidence
  for (const ev of result.qaEvidence) {
    console.log(`  Evidence ${ev.changeId}: ${ev.verificationStatus}`);
    if (ev.failureClassification) {
      console.log(`    Classification: ${ev.failureClassification}`);
      console.log(`    Details: ${ev.failureDetails}`);
    }
  }
}

export default router;
