/**
 * Block-based Contract Editor
 *
 * Key Design Principles:
 * 1. LLM produces ONE replace_block for party definition, not 6 field changes
 * 2. Locator works at paragraph level with normalized text, then maps to runs
 * 3. Handles run fragmentation by joining runs before matching
 * 4. Verification logs required_tokens and missing_tokens explicitly
 */

import OpenAI from 'openai';
import AdmZip from 'adm-zip';
import { parseStringPromise, Builder } from 'xml2js';

// ============================================================
// TYPES
// ============================================================

interface BlockReplacement {
  blockType: 'party_definition';
  anchorText: string;        // Text to find (from original document)
  replacementText: string;   // New text to replace with
  preservedPhrases: string[]; // Legal phrases to preserve
  requiredTokens: string[];  // Tokens that must exist after replacement
}

interface EditResult {
  success: boolean;
  blockType: string;
  anchorFound: boolean;
  paragraphIndex?: number;
  beforeText: string;
  afterText: string;
  appliedAt?: string;
  error?: string;
}

interface VerificationResult {
  passed: boolean;
  requiredTokens: string[];
  foundTokens: string[];
  missingTokens: string[];
  details: string;
}

interface BlockEditResponse {
  success: boolean;
  modifiedBuffer?: Buffer;
  editResult: EditResult;
  verification: VerificationResult;
  logs: string[];
}

// ============================================================
// MAIN EDITOR CLASS
// ============================================================

export class BlockContractEditor {
  private openai: OpenAI;
  private model: string;
  private logs: string[] = [];

  constructor(apiKey: string, model: string = 'gpt-4-turbo-preview') {
    this.openai = new OpenAI({ apiKey });
    this.model = model;
  }

  private log(message: string): void {
    const timestamp = new Date().toISOString();
    const logLine = `[${timestamp}] ${message}`;
    this.logs.push(logLine);
    console.log(logLine);
  }

  /**
   * Main entry point
   */
  async edit(documentBuffer: Buffer, instruction: string): Promise<BlockEditResponse> {
    this.logs = [];
    this.log('=== Block Editor Started ===');

    try {
      // Step 1: Parse DOCX
      this.log('Step 1: Parsing DOCX');
      const zip = new AdmZip(documentBuffer);
      const documentEntry = zip.getEntry('word/document.xml');
      if (!documentEntry) {
        return this.fail('Invalid DOCX: word/document.xml not found');
      }

      const xmlString = documentEntry.getData().toString('utf-8');
      const xml = await parseStringPromise(xmlString, {
        explicitArray: true,
        preserveChildrenOrder: true
      });

      // Step 2: Extract paragraphs with their text
      this.log('Step 2: Extracting paragraphs');
      const paragraphs = this.extractParagraphs(xml);
      this.log(`Found ${paragraphs.length} paragraphs`);

      // Build full document text for LLM
      const fullText = paragraphs.map(p => p.text).join('\n');

      // Step 3: Get block replacement from OpenAI
      this.log('Step 3: Getting block replacement from OpenAI');
      const blockReplacement = await this.getBlockReplacement(fullText, instruction);

      if (!blockReplacement) {
        return this.fail('OpenAI could not identify party definition block');
      }

      this.log(`Block type: ${blockReplacement.blockType}`);
      this.log(`Anchor text (first 100 chars): ${blockReplacement.anchorText.substring(0, 100)}...`);
      this.log(`Required tokens: ${blockReplacement.requiredTokens.join(', ')}`);

      // Step 4: Locate and replace the block
      this.log('Step 4: Locating anchor in paragraphs');
      const editResult = this.locateAndReplace(xml, paragraphs, blockReplacement);

      if (!editResult.success) {
        return {
          success: false,
          editResult,
          verification: {
            passed: false,
            requiredTokens: blockReplacement.requiredTokens,
            foundTokens: [],
            missingTokens: blockReplacement.requiredTokens,
            details: `Edit failed: ${editResult.error}`
          },
          logs: this.logs
        };
      }

      this.log(`Edit applied at paragraph ${editResult.paragraphIndex}`);
      this.log(`BEFORE: ${editResult.beforeText.substring(0, 150)}...`);
      this.log(`AFTER: ${editResult.afterText.substring(0, 150)}...`);

      // Step 5: Save modified DOCX
      this.log('Step 5: Saving modified DOCX');
      const builder = new Builder({
        renderOpts: { pretty: false },
        xmldec: { version: '1.0', encoding: 'UTF-8', standalone: true }
      });
      const modifiedXml = builder.buildObject(xml);
      zip.updateFile('word/document.xml', Buffer.from(modifiedXml, 'utf-8'));
      const modifiedBuffer = zip.toBuffer();

      // Step 6: Verify the changes
      this.log('Step 6: Verifying changes');
      const verification = await this.verifyChanges(
        modifiedBuffer,
        blockReplacement.requiredTokens
      );

      this.log(`Verification: ${verification.passed ? 'PASSED' : 'FAILED'}`);
      this.log(`Required tokens: [${verification.requiredTokens.join(', ')}]`);
      this.log(`Found tokens: [${verification.foundTokens.join(', ')}]`);
      this.log(`Missing tokens: [${verification.missingTokens.join(', ')}]`);
      this.log(`Details: ${verification.details}`);

      return {
        success: verification.passed,
        modifiedBuffer: verification.passed ? modifiedBuffer : undefined,
        editResult,
        verification,
        logs: this.logs
      };

    } catch (error) {
      this.log(`Error: ${error}`);
      return this.fail(error instanceof Error ? error.message : 'Unknown error');
    }
  }

  /**
   * Create a failure response
   */
  private fail(error: string): BlockEditResponse {
    return {
      success: false,
      editResult: {
        success: false,
        blockType: 'party_definition',
        anchorFound: false,
        beforeText: '',
        afterText: '',
        error
      },
      verification: {
        passed: false,
        requiredTokens: [],
        foundTokens: [],
        missingTokens: [],
        details: `Failed: ${error}`
      },
      logs: this.logs
    };
  }

  /**
   * Extract paragraphs from XML with their joined text
   */
  private extractParagraphs(xml: unknown): Array<{
    index: number;
    text: string;
    normalizedText: string;
    element: unknown;
    path: string[];
  }> {
    const paragraphs: Array<{
      index: number;
      text: string;
      normalizedText: string;
      element: unknown;
      path: string[];
    }> = [];

    let index = 0;

    const walk = (node: unknown, path: string[]) => {
      if (!node || typeof node !== 'object') return;

      if (Array.isArray(node)) {
        node.forEach((item, i) => walk(item, [...path, `[${i}]`]));
        return;
      }

      const obj = node as Record<string, unknown>;

      // Check if this is a paragraph (w:p)
      for (const key of Object.keys(obj)) {
        if (key === 'w:p' || key.endsWith(':p')) {
          const pElements = obj[key] as unknown[];
          pElements.forEach((pElement, pIndex) => {
            const text = this.extractTextFromParagraph(pElement);
            if (text.trim()) {
              paragraphs.push({
                index: index++,
                text,
                normalizedText: this.normalizeText(text),
                element: pElement,
                path: [...path, key, `[${pIndex}]`]
              });
            }
          });
        } else if (key !== '$' && key !== '_') {
          walk(obj[key], [...path, key]);
        }
      }
    };

    walk(xml, []);
    return paragraphs;
  }

  /**
   * Extract text from a paragraph by joining all runs
   */
  private extractTextFromParagraph(pElement: unknown): string {
    const texts: string[] = [];

    const extractText = (node: unknown) => {
      if (!node || typeof node !== 'object') return;

      if (Array.isArray(node)) {
        node.forEach(item => extractText(item));
        return;
      }

      const obj = node as Record<string, unknown>;

      // Get text from w:t elements
      for (const key of Object.keys(obj)) {
        if (key === 'w:t' || key.endsWith(':t')) {
          const tElements = obj[key] as unknown[];
          for (const t of tElements) {
            if (typeof t === 'string') {
              texts.push(t);
            } else if (t && typeof t === 'object') {
              const tObj = t as Record<string, unknown>;
              if (tObj['_']) {
                texts.push(tObj['_'] as string);
              }
            }
          }
        } else if (key !== '$') {
          extractText(obj[key]);
        }
      }
    };

    extractText(pElement);
    return texts.join('');
  }

  /**
   * Normalize text for matching
   */
  private normalizeText(text: string): string {
    return text.toLowerCase().replace(/\s+/g, ' ').trim();
  }

  /**
   * Get block replacement specification from OpenAI
   */
  private async getBlockReplacement(
    documentText: string,
    instruction: string
  ): Promise<BlockReplacement | null> {
    const prompt = `You are editing a partnership/contract agreement. The user wants to update party information.

DOCUMENT TEXT:
${documentText.substring(0, 10000)}

USER INSTRUCTION:
${instruction}

YOUR TASK:
1. Find the paragraph that defines the SECOND party (을, "Company", or partner) - usually contains company name, representative, contact info, address
2. Create ONE block replacement that replaces the ENTIRE party definition paragraph
3. PRESERVE legal connecting phrases like:
   - "(hereinafter referred to as the "Company")"
   - "having its principal place of business at"
   - Contract party designations

IMPORTANT RULES:
- Do NOT create multiple field changes - create ONE block replacement
- The anchor_text must be the EXACT text from the document (copy verbatim)
- The replacement_text should have the same structure but with new information
- Include ALL new information (company, person, title, email, phone, address) in the replacement block

Respond with JSON:
{
  "blockType": "party_definition",
  "anchorText": "EXACT text from document to find (the entire party definition sentence/paragraph)",
  "replacementText": "New text with updated information, preserving legal phrases",
  "preservedPhrases": ["list of legal phrases that were preserved"],
  "requiredTokens": ["key words that MUST appear in final document to verify success - include new company name, person name, email, key address words"]
}

Example structure:
- anchorText: "ABC Company, represented by John Doe, (hereinafter referred to as the \"Company\") having its principal place of business at 123 Old Street, Old City"
- replacementText: "XYZ Corp, represented by Jane Smith, (hereinafter referred to as the \"Company\") having its principal place of business at 456 New Street, New City"
- requiredTokens: ["XYZ Corp", "Jane Smith", "456 New Street"]

Respond ONLY with valid JSON.`;

    try {
      const response = await this.openai.chat.completions.create({
        model: this.model,
        messages: [{ role: 'user', content: prompt }],
        temperature: 0,
        response_format: { type: 'json_object' }
      });

      const content = response.choices[0]?.message?.content;
      if (!content) {
        this.log('OpenAI returned empty response');
        return null;
      }

      const result = JSON.parse(content);
      return {
        blockType: result.blockType || 'party_definition',
        anchorText: result.anchorText || '',
        replacementText: result.replacementText || '',
        preservedPhrases: result.preservedPhrases || [],
        requiredTokens: result.requiredTokens || []
      };
    } catch (error) {
      this.log(`OpenAI error: ${error}`);
      return null;
    }
  }

  /**
   * Locate anchor in paragraphs and replace
   * Uses normalized text matching at paragraph level, then maps to runs
   */
  private locateAndReplace(
    xml: unknown,
    paragraphs: Array<{ index: number; text: string; normalizedText: string; element: unknown; path: string[] }>,
    replacement: BlockReplacement
  ): EditResult {
    const normalizedAnchor = this.normalizeText(replacement.anchorText);
    this.log(`Looking for normalized anchor: "${normalizedAnchor.substring(0, 80)}..."`);

    // Find paragraph containing the anchor
    let targetParagraph: typeof paragraphs[0] | null = null;

    for (const para of paragraphs) {
      if (para.normalizedText.includes(normalizedAnchor)) {
        targetParagraph = para;
        this.log(`Found exact match in paragraph ${para.index}`);
        break;
      }
    }

    // If no exact match, try fuzzy matching
    if (!targetParagraph) {
      this.log('No exact match, trying fuzzy match...');

      // Try matching with first 50 chars of anchor
      const shortAnchor = normalizedAnchor.substring(0, 50);
      for (const para of paragraphs) {
        if (para.normalizedText.includes(shortAnchor)) {
          targetParagraph = para;
          this.log(`Found fuzzy match in paragraph ${para.index}`);
          break;
        }
      }
    }

    // If still no match, try matching key identifying phrases
    if (!targetParagraph) {
      this.log('No fuzzy match, trying key phrase match...');

      // Extract key phrases from anchor (company names, etc.)
      const keyPhrases = this.extractKeyPhrases(replacement.anchorText);
      this.log(`Key phrases: ${keyPhrases.join(', ')}`);

      for (const para of paragraphs) {
        const matchCount = keyPhrases.filter(phrase =>
          para.normalizedText.includes(this.normalizeText(phrase))
        ).length;

        if (matchCount >= Math.ceil(keyPhrases.length / 2)) {
          targetParagraph = para;
          this.log(`Found key phrase match (${matchCount}/${keyPhrases.length}) in paragraph ${para.index}`);
          break;
        }
      }
    }

    if (!targetParagraph) {
      // Log available paragraphs for debugging
      this.log('Available paragraphs:');
      paragraphs.slice(0, 10).forEach(p => {
        this.log(`  [${p.index}]: "${p.text.substring(0, 100)}..."`);
      });

      return {
        success: false,
        blockType: replacement.blockType,
        anchorFound: false,
        beforeText: '',
        afterText: '',
        error: `Anchor text not found in any paragraph. Searched for: "${normalizedAnchor.substring(0, 50)}..."`
      };
    }

    // Replace the entire paragraph content
    const beforeText = targetParagraph.text;
    const afterText = replacement.replacementText;

    // Apply replacement by modifying all runs in the paragraph
    this.replaceAllRunsInParagraph(
      targetParagraph.element as Record<string, unknown>,
      afterText
    );

    return {
      success: true,
      blockType: replacement.blockType,
      anchorFound: true,
      paragraphIndex: targetParagraph.index,
      beforeText,
      afterText,
      appliedAt: `Paragraph ${targetParagraph.index}`
    };
  }

  /**
   * Extract key identifying phrases from anchor text
   */
  private extractKeyPhrases(text: string): string[] {
    const phrases: string[] = [];

    // Extract capitalized words/phrases (likely company names, person names)
    const capitalizedPattern = /[A-Z][a-zA-Z]+(?:\s+[A-Z][a-zA-Z]+)*/g;
    const matches = text.match(capitalizedPattern) || [];
    phrases.push(...matches.filter(m => m.length > 3));

    // Extract email-like patterns
    const emailPattern = /[\w.-]+@[\w.-]+/g;
    const emails = text.match(emailPattern) || [];
    phrases.push(...emails);

    return [...new Set(phrases)].slice(0, 5);
  }

  /**
   * Replace all runs in a paragraph with new text
   */
  private replaceAllRunsInParagraph(
    pElement: Record<string, unknown>,
    newText: string
  ): void {
    // Find the run key (w:r or similar)
    let runKey = 'w:r';
    for (const key of Object.keys(pElement)) {
      if (key === 'w:r' || key.endsWith(':r')) {
        runKey = key;
        break;
      }
    }

    // Find paragraph properties key
    let pPrKey = 'w:pPr';
    let pPr: unknown = null;
    for (const key of Object.keys(pElement)) {
      if (key === 'w:pPr' || key.endsWith(':pPr')) {
        pPrKey = key;
        pPr = pElement[key];
        break;
      }
    }

    // Get first run's properties to preserve formatting
    let runProps: unknown = null;
    const existingRuns = pElement[runKey] as unknown[] | undefined;
    if (existingRuns && existingRuns.length > 0) {
      const firstRun = existingRuns[0] as Record<string, unknown>;
      for (const key of Object.keys(firstRun)) {
        if (key === 'w:rPr' || key.endsWith(':rPr')) {
          runProps = JSON.parse(JSON.stringify(firstRun[key]));
          break;
        }
      }
    }

    // Create new run with the replacement text and yellow highlight
    const newRun: Record<string, unknown> = {};

    // Add run properties with highlight
    const rPr: Record<string, unknown[]> = runProps
      ? JSON.parse(JSON.stringify((runProps as unknown[])[0]))
      : {};
    rPr['w:highlight'] = [{ $: { 'w:val': 'yellow' } }];
    newRun['w:rPr'] = [rPr];

    // Add text element
    newRun['w:t'] = [{ _: newText, $: { 'xml:space': 'preserve' } }];

    // Replace all runs with single new run
    pElement[runKey] = [newRun];

    // Preserve paragraph properties
    if (pPr) {
      pElement[pPrKey] = pPr;
    }
  }

  /**
   * Verify changes by checking required tokens exist in final document
   */
  private async verifyChanges(
    modifiedBuffer: Buffer,
    requiredTokens: string[]
  ): Promise<VerificationResult> {
    try {
      const zip = new AdmZip(modifiedBuffer);
      const documentEntry = zip.getEntry('word/document.xml');
      if (!documentEntry) {
        return {
          passed: false,
          requiredTokens,
          foundTokens: [],
          missingTokens: requiredTokens,
          details: 'Could not read modified document for verification'
        };
      }

      const xmlString = documentEntry.getData().toString('utf-8');

      // Check for [object Object] contamination
      if (xmlString.includes('[object Object]')) {
        return {
          passed: false,
          requiredTokens,
          foundTokens: [],
          missingTokens: requiredTokens,
          details: 'Document contains [object Object] contamination - XML serialization failed'
        };
      }

      // Parse and extract text
      const xml = await parseStringPromise(xmlString, { explicitArray: true });
      const paragraphs = this.extractParagraphs(xml);
      const fullText = paragraphs.map(p => p.text).join(' ');
      const normalizedFullText = this.normalizeText(fullText);

      // Check each required token
      const foundTokens: string[] = [];
      const missingTokens: string[] = [];

      for (const token of requiredTokens) {
        const normalizedToken = this.normalizeText(token);
        if (normalizedFullText.includes(normalizedToken)) {
          foundTokens.push(token);
        } else {
          missingTokens.push(token);
        }
      }

      const passed = missingTokens.length === 0;

      let details: string;
      if (passed) {
        details = `All ${requiredTokens.length} required tokens found in document`;
      } else {
        details = `Missing ${missingTokens.length}/${requiredTokens.length} tokens: [${missingTokens.join(', ')}]`;
      }

      return {
        passed,
        requiredTokens,
        foundTokens,
        missingTokens,
        details
      };

    } catch (error) {
      return {
        passed: false,
        requiredTokens,
        foundTokens: [],
        missingTokens: requiredTokens,
        details: `Verification error: ${error instanceof Error ? error.message : 'Unknown error'}`
      };
    }
  }
}
