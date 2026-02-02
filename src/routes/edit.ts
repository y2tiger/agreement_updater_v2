/**
 * Edit API Routes
 *
 * Simplified contract editing using find/replace approach.
 */

import { Router, Request, Response } from 'express';
import multer from 'multer';
import { v4 as uuidv4 } from 'uuid';
import * as fs from 'fs';
import * as path from 'path';
import { SimpleContractEditor } from '../simple-editor';

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
    console.log(`[${requestId}] Instruction: ${instruction}`);

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

    // Use simplified editor
    const editor = new SimpleContractEditor(
      openaiApiKey,
      process.env.OPENAI_MODEL || 'gpt-4-turbo-preview'
    );

    const result = await editor.edit(documentBuffer, instruction);

    // Clean up uploaded file
    fs.unlinkSync(uploadedFilePath);
    uploadedFilePath = undefined;

    // Build response
    const response: any = {
      success: result.success,
      changes: result.changes.map(c => ({
        description: c.description,
        find: c.find,
        replace: c.replace,
        status: c.applied ? 'APPLIED' : 'NOT_FOUND',
        count: c.count
      })),
      error: result.error
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

      console.log(`[${requestId}] Edit completed successfully in ${Date.now() - startTime}ms`);
      console.log(`[${requestId}] Applied ${result.changes.filter(c => c.applied).length}/${result.changes.length} changes`);
    } else {
      console.log(`[${requestId}] Edit failed: ${result.error}`);
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

export default router;
