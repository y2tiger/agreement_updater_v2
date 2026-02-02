/**
 * Edit API Routes
 *
 * Uses block-based editing approach:
 * - ONE replace_block for party definition (not 6 field changes)
 * - Paragraph-level normalized text matching
 * - Explicit verification with required_tokens logging
 */

import { Router, Request, Response } from 'express';
import multer from 'multer';
import { v4 as uuidv4 } from 'uuid';
import * as fs from 'fs';
import * as path from 'path';
import { BlockContractEditor } from '../block-editor';

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
}, 60000);

/**
 * POST /api/edit
 *
 * Upload a DOCX file and apply block-based edits
 */
router.post('/edit', upload.single('file'), async (req: Request, res: Response) => {
  const requestId = uuidv4();
  const startTime = Date.now();

  console.log(`\n${'='.repeat(60)}`);
  console.log(`[${requestId}] Edit request received`);

  let uploadedFilePath: string | undefined;

  try {
    // Validate request
    if (!req.file) {
      return res.status(400).json({
        success: false,
        error: 'No file uploaded',
      });
    }

    const instruction = req.body.instruction as string;
    if (!instruction || instruction.trim().length === 0) {
      fs.unlinkSync(req.file.path);
      return res.status(400).json({
        success: false,
        error: 'No instruction provided',
      });
    }

    uploadedFilePath = req.file.path;
    const fileName = req.file.originalname;

    console.log(`[${requestId}] File: ${fileName}`);
    console.log(`[${requestId}] Instruction: ${instruction.substring(0, 200)}...`);

    // Check for OpenAI API key
    const openaiApiKey = process.env.OPENAI_API_KEY;
    if (!openaiApiKey) {
      fs.unlinkSync(uploadedFilePath);
      return res.status(500).json({
        success: false,
        error: 'OpenAI API key not configured',
      });
    }

    // Read the uploaded file
    const documentBuffer = fs.readFileSync(uploadedFilePath);

    // Use block editor
    const editor = new BlockContractEditor(
      openaiApiKey,
      process.env.OPENAI_MODEL || 'gpt-4-turbo-preview'
    );

    const result = await editor.edit(documentBuffer, instruction);

    // Clean up uploaded file
    fs.unlinkSync(uploadedFilePath);
    uploadedFilePath = undefined;

    // Log results
    console.log(`[${requestId}] Pipeline Result:`);
    console.log(`  Success: ${result.success}`);
    console.log(`  Edit Result:`);
    console.log(`    - Block Type: ${result.editResult.blockType}`);
    console.log(`    - Anchor Found: ${result.editResult.anchorFound}`);
    console.log(`    - Applied At: ${result.editResult.appliedAt || 'N/A'}`);
    if (result.editResult.error) {
      console.log(`    - Error: ${result.editResult.error}`);
    }
    console.log(`  Verification:`);
    console.log(`    - Passed: ${result.verification.passed}`);
    console.log(`    - Required Tokens: [${result.verification.requiredTokens.join(', ')}]`);
    console.log(`    - Found Tokens: [${result.verification.foundTokens.join(', ')}]`);
    console.log(`    - Missing Tokens: [${result.verification.missingTokens.join(', ')}]`);
    console.log(`    - Details: ${result.verification.details}`);
    console.log(`  Duration: ${Date.now() - startTime}ms`);
    console.log('='.repeat(60));

    // Build response
    const response: Record<string, unknown> = {
      success: result.success,
      editResult: {
        blockType: result.editResult.blockType,
        anchorFound: result.editResult.anchorFound,
        appliedAt: result.editResult.appliedAt,
        beforeText: result.editResult.beforeText.substring(0, 300) + (result.editResult.beforeText.length > 300 ? '...' : ''),
        afterText: result.editResult.afterText.substring(0, 300) + (result.editResult.afterText.length > 300 ? '...' : ''),
        error: result.editResult.error
      },
      verification: result.verification,
      error: result.success ? undefined : (result.editResult.error || result.verification.details)
    };

    if (result.success && result.modifiedBuffer) {
      // Store the modified file for download
      const downloadId = uuidv4();
      const modifiedFileName = `edited_${fileName}`;

      processedFiles.set(downloadId, {
        buffer: result.modifiedBuffer,
        fileName: modifiedFileName,
        expiry: Date.now() + 15 * 60 * 1000, // 15 minutes
      });

      response.downloadUrl = `/api/download/${downloadId}`;
    }

    return res.json(response);

  } catch (error) {
    console.error(`[${requestId}] Error:`, error);

    // Clean up uploaded file on error
    if (uploadedFilePath && fs.existsSync(uploadedFilePath)) {
      fs.unlinkSync(uploadedFilePath);
    }

    return res.status(500).json({
      success: false,
      error: error instanceof Error ? error.message : 'Internal server error',
    });
  }
});

/**
 * GET /api/download/:id
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

  processedFiles.delete(id);
});

/**
 * GET /api/health
 */
router.get('/health', (req: Request, res: Response) => {
  res.json({
    status: 'ok',
    timestamp: new Date().toISOString(),
    openaiConfigured: !!process.env.OPENAI_API_KEY,
    approach: 'block-based-editing'
  });
});

export default router;
