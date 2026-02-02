/**
 * Simplified Contract Editor
 *
 * Simple find/replace approach that actually works:
 * 1. Parse DOCX and extract text
 * 2. Use OpenAI to identify specific find/replace pairs
 * 3. Apply replacements directly to run text
 * 4. Highlight modified text
 * 5. Save
 */

import OpenAI from 'openai';
import AdmZip from 'adm-zip';
import { parseStringPromise, Builder } from 'xml2js';

export interface FindReplacePair {
  find: string;
  replace: string;
  description: string;
}

export interface SimpleEditResult {
  success: boolean;
  modifiedBuffer?: Buffer;
  changes: Array<{
    find: string;
    replace: string;
    description: string;
    applied: boolean;
    count: number;
  }>;
  error?: string;
}

/**
 * Simple contract editor using find/replace
 */
export class SimpleContractEditor {
  private openai: OpenAI;
  private model: string;

  constructor(apiKey: string, model: string = 'gpt-4-turbo-preview') {
    this.openai = new OpenAI({ apiKey });
    this.model = model;
  }

  /**
   * Main entry point
   */
  async edit(documentBuffer: Buffer, instruction: string): Promise<SimpleEditResult> {
    try {
      // Step 1: Parse DOCX
      const zip = new AdmZip(documentBuffer);
      const documentEntry = zip.getEntry('word/document.xml');
      if (!documentEntry) {
        return { success: false, changes: [], error: 'Invalid DOCX file' };
      }

      const xmlString = documentEntry.getData().toString('utf-8');
      const xml = await parseStringPromise(xmlString, {
        explicitArray: true,
        preserveChildrenOrder: true
      });

      // Step 2: Extract plain text for analysis
      const plainText = this.extractText(xml);
      console.log('Extracted text length:', plainText.length);

      // Step 3: Get find/replace pairs from OpenAI
      const pairs = await this.getFindReplacePairs(plainText, instruction);
      console.log('Find/replace pairs:', pairs.length);

      if (pairs.length === 0) {
        return {
          success: false,
          changes: [],
          error: 'Could not identify any changes to make'
        };
      }

      // Step 4: Apply replacements
      const results = this.applyReplacements(xml, pairs);

      // Check if any changes were made
      const appliedCount = results.filter(r => r.applied).length;
      if (appliedCount === 0) {
        return {
          success: false,
          changes: results,
          error: 'No text matches found in document'
        };
      }

      // Step 5: Save modified DOCX
      const builder = new Builder({
        renderOpts: { pretty: false },
        xmldec: { version: '1.0', encoding: 'UTF-8', standalone: true }
      });
      const modifiedXml = builder.buildObject(xml);

      zip.updateFile('word/document.xml', Buffer.from(modifiedXml, 'utf-8'));
      const modifiedBuffer = zip.toBuffer();

      return {
        success: true,
        modifiedBuffer,
        changes: results
      };

    } catch (error) {
      console.error('Edit error:', error);
      return {
        success: false,
        changes: [],
        error: error instanceof Error ? error.message : 'Unknown error'
      };
    }
  }

  /**
   * Extract plain text from parsed XML
   */
  private extractText(xml: unknown): string {
    const texts: string[] = [];
    this.walkXml(xml, (node) => {
      if (node && typeof node === 'object') {
        const n = node as Record<string, unknown>;
        // Get text from w:t elements
        if (n['w:t']) {
          const tElements = n['w:t'] as unknown[];
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
        }
      }
    });
    return texts.join('');
  }

  /**
   * Walk XML tree and call callback for each node
   */
  private walkXml(node: unknown, callback: (node: unknown) => void): void {
    callback(node);
    if (node && typeof node === 'object') {
      if (Array.isArray(node)) {
        for (const item of node) {
          this.walkXml(item, callback);
        }
      } else {
        for (const key of Object.keys(node as Record<string, unknown>)) {
          this.walkXml((node as Record<string, unknown>)[key], callback);
        }
      }
    }
  }

  /**
   * Get find/replace pairs from OpenAI
   */
  private async getFindReplacePairs(documentText: string, instruction: string): Promise<FindReplacePair[]> {
    const prompt = `You are helping edit a contract document. The user wants to make changes.

DOCUMENT TEXT (excerpt):
${documentText.substring(0, 8000)}

USER INSTRUCTION:
${instruction}

Your task: Identify specific text strings that need to be found and replaced.

IMPORTANT RULES:
1. The "find" text MUST be an EXACT substring that exists in the document
2. Keep "find" text SHORT (1-5 words) for reliable matching
3. Be specific - find the exact company name, person name, email, phone, address that needs to change
4. Do NOT include surrounding context in "find" - just the specific text to replace

Respond with JSON:
{
  "pairs": [
    {
      "find": "exact text to find (SHORT, 1-5 words)",
      "replace": "new text",
      "description": "what this changes"
    }
  ]
}

Example - if changing company name from "ABC Corp" to "XYZ Inc":
{
  "pairs": [
    {"find": "ABC Corp", "replace": "XYZ Inc", "description": "Company name"}
  ]
}

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
        return [];
      }

      const result = JSON.parse(content);
      return result.pairs || [];
    } catch (error) {
      console.error('OpenAI error:', error);
      return [];
    }
  }

  /**
   * Apply find/replace to the XML
   */
  private applyReplacements(
    xml: unknown,
    pairs: FindReplacePair[]
  ): Array<{ find: string; replace: string; description: string; applied: boolean; count: number }> {
    const results: Array<{ find: string; replace: string; description: string; applied: boolean; count: number }> = [];

    for (const pair of pairs) {
      let count = 0;

      // Walk through all text elements and replace
      this.walkAndReplace(xml, pair.find, pair.replace, (replaced) => {
        if (replaced) count++;
      });

      results.push({
        find: pair.find,
        replace: pair.replace,
        description: pair.description,
        applied: count > 0,
        count
      });

      console.log(`Replace "${pair.find}" -> "${pair.replace}": ${count} occurrences`);
    }

    return results;
  }

  /**
   * Walk XML and replace text, adding highlight
   */
  private walkAndReplace(
    node: unknown,
    find: string,
    replace: string,
    onReplace: (replaced: boolean) => void
  ): void {
    if (!node || typeof node !== 'object') return;

    if (Array.isArray(node)) {
      for (const item of node) {
        this.walkAndReplace(item, find, replace, onReplace);
      }
      return;
    }

    const obj = node as Record<string, unknown>;

    // Check if this is a run (w:r) element with text
    if (obj['w:t']) {
      const tElements = obj['w:t'] as unknown[];

      for (let i = 0; i < tElements.length; i++) {
        const t = tElements[i];
        let text: string | undefined;
        let isObject = false;

        if (typeof t === 'string') {
          text = t;
        } else if (t && typeof t === 'object') {
          const tObj = t as Record<string, unknown>;
          if (tObj['_']) {
            text = tObj['_'] as string;
            isObject = true;
          }
        }

        if (text && text.includes(find)) {
          const newText = text.split(find).join(replace);

          if (isObject) {
            (tElements[i] as Record<string, unknown>)['_'] = newText;
          } else {
            tElements[i] = { _: newText, $: { 'xml:space': 'preserve' } };
          }

          // Add yellow highlight to the run
          this.addHighlight(obj);
          onReplace(true);
        }
      }
    }

    // Recurse into children
    for (const key of Object.keys(obj)) {
      if (key !== '$') {
        this.walkAndReplace(obj[key], find, replace, onReplace);
      }
    }
  }

  /**
   * Add yellow highlight to a run element
   */
  private addHighlight(runElement: Record<string, unknown>): void {
    // Get or create run properties
    if (!runElement['w:rPr']) {
      runElement['w:rPr'] = [{}];
    }

    const rPr = (runElement['w:rPr'] as unknown[])[0] as Record<string, unknown>;

    // Add highlight
    rPr['w:highlight'] = [{ $: { 'w:val': 'yellow' } }];
  }
}
