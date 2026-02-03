/**
 * Edit API Routes
 *
 * General-Purpose Contract Editing Engine
 *
 * This system converts natural language change requests into
 * structurally safe DOCX modifications.
 *
 * NOT domain-specific. No assumptions about company/address/contact.
 */

import { Router, Request, Response } from 'express';
import multer from 'multer';
import { v4 as uuidv4 } from 'uuid';
import * as fs from 'fs';
import * as path from 'path';
import { editContract } from '../contract-engine';

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
 * General-purpose contract editing endpoint
 *
 * Required output for every request:
 * 1) Change Items (request decomposition)
 * 2) Candidate edit locations + rationale
 * 3) Final Edit Specs
 * 4) BEFORE / AFTER evidence
 * 5) Verification result (PASS / FAIL + reason)
 */
router.post('/edit', upload.single('file'), async (req: Request, res: Response) => {
  const requestId = uuidv4();
  const startTime = Date.now();

  console.log(`\n${'='.repeat(70)}`);
  console.log(`[${requestId}] CONTRACT EDIT REQUEST`);
  console.log('='.repeat(70));

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

    console.log(`File: ${fileName}`);
    console.log(`User Request: ${instruction}`);
    console.log('-'.repeat(70));

    // Check for OpenAI API key
    if (!process.env.OPENAI_API_KEY) {
      fs.unlinkSync(uploadedFilePath);
      return res.status(500).json({
        success: false,
        error: 'OpenAI API key not configured',
      });
    }

    // Read the uploaded file
    const documentBuffer = fs.readFileSync(uploadedFilePath);

    // Execute the generalized contract editing pipeline
    const result = await editContract(documentBuffer, instruction);

    // Clean up uploaded file
    fs.unlinkSync(uploadedFilePath);
    uploadedFilePath = undefined;

    // ========================
    // OUTPUT: Required Logging
    // ========================

    // 1) Change Items
    console.log('\n[1] CHANGE ITEMS (Request Decomposition):');
    for (const item of (result.changeItems || [])) {
      console.log(`  - ID: ${item?.id ?? 'N/A'}`);
      console.log(`    Intent: ${item?.intent ?? 'N/A'}`);
      console.log(`    Fragment: "${item?.userRequestFragment ?? 'N/A'}"`);
      if (item?.ambiguityNote) {
        console.log(`    ⚠️ Ambiguity: ${item.ambiguityNote}`);
      }
    }

    // 2) Candidate Locations
    console.log('\n[2] CANDIDATE LOCATIONS:');
    for (const item of (result.changeItems || [])) {
      console.log(`  Change Item: ${item?.id ?? 'N/A'}`);
      for (const candidate of (item?.candidates || [])) {
        const excerpt = candidate?.excerptFromDocument ?? '';
        console.log(`    - Location: ${candidate?.locationDescription ?? 'N/A'}`);
        console.log(`      Excerpt: "${excerpt.substring(0, 80)}${excerpt.length > 80 ? '...' : ''}"`);
        console.log(`      Rationale: ${candidate?.rationale ?? 'N/A'}`);
        console.log(`      Confidence: ${candidate?.confidence != null ? (candidate.confidence * 100).toFixed(0) : 'N/A'}%`);
      }
    }

    // 3) Edit Specs
    console.log('\n[3] EDIT SPECIFICATIONS:');
    for (const spec of (result.editSpecs || [])) {
      const anchor = spec?.anchorText ?? '';
      const before = spec?.beforeText ?? '';
      const after = spec?.afterText ?? '';
      console.log(`  - Spec for: ${spec?.changeItemId ?? 'N/A'}`);
      console.log(`    Target Unit: ${spec?.targetUnit ?? 'N/A'}`);
      console.log(`    Edit Type: ${spec?.editType ?? 'N/A'}`);
      console.log(`    Anchor: "${anchor.substring(0, 60)}${anchor.length > 60 ? '...' : ''}"`);
      console.log(`    Before: "${before.substring(0, 60)}${before.length > 60 ? '...' : ''}"`);
      console.log(`    After: "${after.substring(0, 60)}${after.length > 60 ? '...' : ''}"`);
    }

    // 4) BEFORE / AFTER Evidence
    console.log('\n[4] BEFORE / AFTER EVIDENCE:');
    for (const exec of (result.executions || [])) {
      const actualBefore = exec?.actualBefore ?? '';
      const actualAfter = exec?.actualAfter ?? '';
      console.log(`  - Spec: ${exec?.editSpecId ?? 'N/A'}`);
      console.log(`    Anchor Found: ${exec?.anchorFound ?? 'N/A'}`);
      console.log(`    Applied: ${exec?.applied ?? 'N/A'}`);
      if (exec?.blockReason) {
        console.log(`    ❌ Block Reason: ${exec.blockReason}`);
      }
      console.log(`    BEFORE: "${actualBefore.substring(0, 100)}${actualBefore.length > 100 ? '...' : ''}"`);
      console.log(`    AFTER:  "${actualAfter.substring(0, 100)}${actualAfter.length > 100 ? '...' : ''}"`);
    }

    // 5) Verification Result
    console.log('\n[5] VERIFICATION RESULT:');
    console.log(`  Status: ${result.verification?.status ?? 'N/A'}`);
    console.log(`  Summary: ${result.verification?.summary ?? 'N/A'}`);
    console.log(`  Document Integrity:`);
    console.log(`    - Valid: ${result.verification?.documentIntegrity?.valid ?? 'N/A'}`);
    console.log(`    - Can Open: ${result.verification?.documentIntegrity?.canOpen ?? 'N/A'}`);
    console.log(`    - Structure Preserved: ${result.verification?.documentIntegrity?.structurePreserved ?? 'N/A'}`);
    console.log('  Spec Results:');
    for (const specResult of (result.verification?.specResults || [])) {
      console.log(`    - ${specResult?.editSpecId ?? 'N/A'}: ${specResult?.status ?? 'N/A'}`);
      console.log(`      after_text found: ${specResult?.afterTextFound ?? 'N/A'}`);
      console.log(`      location correct: ${specResult?.locationCorrect ?? 'N/A'}`);
      console.log(`      match: ${specResult?.evidence?.matchPercentage ?? 'N/A'}%`);
      if (specResult?.failureReason) {
        console.log(`      ❌ Reason: ${specResult.failureReason}`);
      }
    }

    // Timing
    console.log('\n[TIMING]:');
    console.log(`  Interpret: ${result.timing?.interpretMs ?? 'N/A'}ms`);
    console.log(`  Spec Gen:  ${result.timing?.specGenMs ?? 'N/A'}ms`);
    console.log(`  Execute:   ${result.timing?.executeMs ?? 'N/A'}ms`);
    console.log(`  Verify:    ${result.timing?.verifyMs ?? 'N/A'}ms`);
    console.log(`  TOTAL:     ${result.timing?.totalMs ?? 'N/A'}ms`);
    console.log('='.repeat(70));

    // Build API response
    const success = result.verification?.status === 'PASS';
    const response: Record<string, unknown> = {
      success,
      userRequest: result.userRequest,
      changeItems: result.changeItems || [],
      editSpecs: (result.editSpecs || []).map(spec => {
        const anchorText = spec?.anchorText ?? '';
        return {
          changeItemId: spec?.changeItemId,
          targetUnit: spec?.targetUnit,
          editType: spec?.editType,
          anchorText: anchorText.substring(0, 200) + (anchorText.length > 200 ? '...' : ''),
          beforeText: spec?.beforeText ?? '',
          afterText: spec?.afterText ?? '',
          constraints: spec?.constraints ?? [],
        };
      }),
      executions: (result.executions || []).map(exec => {
        const actualBefore = exec?.actualBefore ?? '';
        const actualAfter = exec?.actualAfter ?? '';
        return {
          editSpecId: exec?.editSpecId,
          anchorFound: exec?.anchorFound,
          applied: exec?.applied,
          blockReason: exec?.blockReason,
          beforeText: actualBefore.substring(0, 300) + (actualBefore.length > 300 ? '...' : ''),
          afterText: actualAfter.substring(0, 300) + (actualAfter.length > 300 ? '...' : ''),
          locationIndex: exec?.locationIndex,
        };
      }),
      verification: {
        status: result.verification?.status ?? 'FAIL',
        summary: result.verification?.summary ?? 'Unknown error occurred',
        documentIntegrity: result.verification?.documentIntegrity ?? { valid: false, canOpen: false, structurePreserved: false },
        specResults: result.verification?.specResults ?? [],
      },
      timing: result.timing ?? {},
    };

    if (success && result.modifiedDocument) {
      // Store the modified file for download
      const downloadId = uuidv4();
      const modifiedFileName = `edited_${fileName}`;

      processedFiles.set(downloadId, {
        buffer: result.modifiedDocument,
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
    approach: 'general-purpose-contract-engine',
  });
});

export default router;
