/**
 * Agreement Updater v2 - Main Entry Point
 *
 * Contract editing system using OpenAI for planning and safe DOCX manipulation.
 *
 * DEPLOYMENT: Render
 * - Stateless process
 * - All files processed in /tmp directories
 * - Any failure returns clear error response
 * - "Edit completed successfully" only after full verification PASS
 */

import express from 'express';
import cors from 'cors';
import dotenv from 'dotenv';
import * as fs from 'fs';
import * as path from 'path';
import editRoutes from './routes/edit';

// Load environment variables
dotenv.config();

const app = express();
const PORT = process.env.PORT || 3000;

// Middleware
app.use(cors());
app.use(express.json());

// Serve static files from public directory
app.use(express.static(path.join(__dirname, '../public')));

// Ensure /tmp directory exists
const tmpDir = '/tmp/contract-edits';
if (!fs.existsSync(tmpDir)) {
  fs.mkdirSync(tmpDir, { recursive: true });
}

// Routes
app.use('/api', editRoutes);

// API info endpoint
app.get('/api/info', (req, res) => {
  res.json({
    name: 'Agreement Updater v2',
    version: '1.0.0',
    description: 'Contract editing system using OpenAI for planning and safe DOCX manipulation',
    endpoints: {
      'POST /api/edit': 'Upload DOCX and apply natural language edits',
      'GET /api/download/:id': 'Download modified DOCX',
      'GET /api/health': 'Health check',
    },
    documentation: {
      system: {
        llmProvider: 'OpenAI (for reasoning, targeting, Edit Spec generation only)',
        deployment: 'Render (stateless, /tmp processing)',
        safety: 'Full verification required before success response',
      },
      pipeline: [
        'PHASE 0: Global Context Scan (OpenAI, read-only)',
        'PHASE 1: Change Decomposition (OpenAI)',
        'PHASE 2: Candidate Targeting (OpenAI)',
        'PHASE 3: Edit Spec Lock (OpenAI)',
        'PHASE 4: Apply (Code only)',
        'PHASE 5: Post-Verification (Guardian)',
      ],
      agents: [
        'Contract Architect: Plans edits using OpenAI',
        'DOCX Locator: Maps targets to DOCX elements',
        'Edit Executor: Applies changes with highlighting',
        'Flow Guardian: Verifies all changes',
        'QA Evidence: Documents BEFORE/AFTER',
      ],
      blockConditions: [
        'Ambiguous target',
        'Verification failure',
        '[object Object] detected',
        'DOCX integrity failure',
      ],
    },
  });
});

// Error handling middleware
app.use((err: Error, req: express.Request, res: express.Response, next: express.NextFunction) => {
  console.error('Unhandled error:', err);
  res.status(500).json({
    success: false,
    blockedReason: 'Internal server error',
    error: process.env.NODE_ENV === 'development' ? err.message : undefined,
  });
});

// Handle 404
app.use((req, res) => {
  res.status(404).json({
    success: false,
    error: 'Not found',
  });
});

// Graceful shutdown
process.on('SIGTERM', () => {
  console.log('SIGTERM received. Shutting down gracefully...');
  process.exit(0);
});

process.on('SIGINT', () => {
  console.log('SIGINT received. Shutting down gracefully...');
  process.exit(0);
});

// Start server
app.listen(PORT, () => {
  console.log(`Agreement Updater v2 running on port ${PORT}`);
  console.log(`OpenAI configured: ${!!process.env.OPENAI_API_KEY}`);
  console.log(`Model: ${process.env.OPENAI_MODEL || 'gpt-4-turbo-preview'}`);
});

export default app;
