/**
 * General-Purpose Contract Editing Engine
 *
 * This system converts natural language change requests into
 * structurally safe DOCX modifications.
 *
 * NOT a domain-specific tool. No assumptions about:
 * - Company names, addresses, contacts
 * - Specific contract types
 * - Particular fields or sections
 *
 * Core questions this system answers:
 * 1. WHERE to change?
 * 2. HOW MUCH to change?
 * 3. Was the change actually applied?
 */

import JSZip from 'jszip';
import { parseStringPromise, Builder } from 'xml2js';
import OpenAI from 'openai';
import {
  ChangeItem,
  EditSpec,
  EditExecution,
  VerificationResult,
  SpecVerification,
  ContractEditResponse,
  CandidateLocation,
} from './types/edit-spec';

const openai = new OpenAI({ apiKey: process.env.OPENAI_API_KEY });

// ========================
// Document Parser (Utility)
// ========================

interface ParsedParagraph {
  index: number;
  text: string;
  normalizedText: string;
  path: string;
  element: any;
}

interface ParsedDocument {
  paragraphs: ParsedParagraph[];
  fullText: string;
  xml: any;
  zip: JSZip;
}

function findKey(obj: any, ...keys: string[]): string | undefined {
  if (!obj || typeof obj !== 'object') return undefined;
  for (const key of keys) {
    if (key in obj) return key;
  }
  return undefined;
}

function extractText(element: any): string {
  if (!element) return '';
  if (typeof element === 'string') return element;
  if (Array.isArray(element)) return element.map(extractText).join('');

  let text = '';
  const tKey = findKey(element, 'w:t', 't');
  if (tKey) {
    const tElements = Array.isArray(element[tKey]) ? element[tKey] : [element[tKey]];
    for (const t of tElements) {
      if (typeof t === 'string') text += t;
      else if (t && t._) text += t._;
    }
  }

  const rKey = findKey(element, 'w:r', 'r');
  if (rKey) {
    const runs = Array.isArray(element[rKey]) ? element[rKey] : [element[rKey]];
    for (const run of runs) {
      text += extractText(run);
    }
  }

  return text;
}

async function parseDocument(buffer: Buffer): Promise<ParsedDocument> {
  const zip = await JSZip.loadAsync(buffer);
  const documentXml = await zip.file('word/document.xml')?.async('string');

  if (!documentXml) {
    throw new Error('Invalid DOCX: word/document.xml not found');
  }

  const xml = await parseStringPromise(documentXml, { explicitArray: false });
  const docKey = findKey(xml, 'w:document', 'document');
  if (!docKey) {
    throw new Error('Invalid DOCX: document element not found');
  }

  const bodyKey = findKey(xml[docKey], 'w:body', 'body');
  if (!bodyKey) {
    throw new Error('Invalid DOCX: body element not found');
  }

  const body = xml[docKey][bodyKey];
  const pKey = findKey(body, 'w:p', 'p');
  const paragraphElements = pKey ? (Array.isArray(body[pKey]) ? body[pKey] : [body[pKey]]) : [];

  const paragraphs: ParsedParagraph[] = [];
  let fullText = '';

  paragraphElements.forEach((p: any, idx: number) => {
    const text = extractText(p);
    paragraphs.push({
      index: idx,
      text,
      normalizedText: text.replace(/\s+/g, ' ').trim().toLowerCase(),
      path: `${docKey}.${bodyKey}.${pKey}[${idx}]`,
      element: p,
    });
    fullText += text + '\n';
  });

  return { paragraphs, fullText, xml, zip };
}

// ========================
// 1. Change Interpreter (LLM)
// ========================

async function interpretChanges(
  document: ParsedDocument,
  userRequest: string
): Promise<ChangeItem[]> {
  const documentContext = document.paragraphs
    .map((p, i) => `[P${i}] ${p.text}`)
    .join('\n');

  const prompt = `You are a document analysis expert. Analyze a change request and identify WHERE in the document changes should be made.

DOCUMENT CONTENT:
${documentContext}

USER REQUEST:
${userRequest}

CRITICAL RULES FOR DECOMPOSITION:

1. BLOCK REPLACEMENT PRINCIPLE:
   - If user provides a BLOCK of related information (name, title, email, phone, address on multiple lines), treat it as ONE replacement unit
   - Look for a SIMILAR BLOCK in the document that contains the same types of information
   - Do NOT split into separate change items for each field

2. PARTY IDENTIFICATION:
   - Korean contracts use "갑" (Party A/First Party) and "을" (Party B/Second Party)
   - English contracts use "Party A/B", "First Party/Second Party", "Licensor/Licensee", etc.
   - When user says "을 회사" or "Party B", find the party definition block for that party

3. STRUCTURAL MATCHING:
   - Match by STRUCTURE, not by individual values
   - If document has: [Company Name] + [Person Name] + [Title] + [Email] + [Phone] + [Address]
   - And user provides the same structure, it's ONE block replacement

4. CONFIDENCE RULES:
   - If you can identify the target party/section clearly: confidence >= 0.9
   - If the document structure matches user's input structure: confidence >= 0.85
   - Only mark ambiguity if there are truly multiple equally valid locations

RESPOND IN JSON:
{
  "changeItems": [
    {
      "id": "change_1",
      "userRequestFragment": "full description of what user wants to change",
      "intent": "replace",
      "candidates": [
        {
          "locationDescription": "description of where in document",
          "excerptFromDocument": "EXACT text from document that will be replaced",
          "rationale": "why this location matches",
          "confidence": 0.95
        }
      ],
      "ambiguityNote": "only if truly ambiguous"
    }
  ]
}

IMPORTANT: If user provides multi-line contact/company information, create ONLY ONE change item that replaces the entire corresponding block in the document.`;

  const response = await openai.chat.completions.create({
    model: 'gpt-4o',
    messages: [{ role: 'user', content: prompt }],
    response_format: { type: 'json_object' },
    temperature: 0.1,
  });

  const result = JSON.parse(response.choices[0].message.content || '{}');
  return result.changeItems || [];
}

// ========================
// 2. Edit Spec Generator (LLM)
// ========================

async function generateEditSpecs(
  document: ParsedDocument,
  changeItems: ChangeItem[],
  userRequest: string
): Promise<EditSpec[]> {
  const documentContext = document.paragraphs
    .map((p, i) => `[P${i}] ${p.text}`)
    .join('\n');

  const prompt = `You are a precise document editor. For each Change Item, generate an exact Edit Specification.

DOCUMENT CONTENT:
${documentContext}

USER REQUEST:
${userRequest}

CHANGE ITEMS:
${JSON.stringify(changeItems, null, 2)}

TASK:
For each Change Item, create ONE Edit Spec:
1. target_unit: paragraph, table_cell, etc.
2. anchor_text: Exact text from document to locate edit position
3. boundary_start/end: Edit boundaries within anchor
4. edit_type: replace, insert_before, insert_after, delete
5. before_text: EXACT text being replaced (copy from document)
6. after_text: New text to insert
7. constraints: What to preserve

CRITICAL RULES:

1. EXACT TEXT MATCHING:
   - before_text MUST be copied EXACTLY from the document
   - Include the FULL text block that needs replacement
   - For multi-line party info, include ALL lines in before_text

2. BLOCK REPLACEMENT:
   - If Change Item represents a block of information (company + contact details)
   - before_text = entire existing block from document
   - after_text = entire new block from user request
   - Preserve the same line break/separator style as original

3. FORMATTING:
   - Match the formatting style of the original document
   - If original uses "E:" for email, keep that prefix
   - If original uses line breaks, keep line breaks

4. DO NOT BLOCK unless truly impossible:
   - If candidate location is found with confidence >= 0.8, proceed
   - Only use edit_type "blocked" if no valid location exists

RESPOND IN JSON:
{
  "editSpecs": [
    {
      "changeItemId": "change_1",
      "targetUnit": "paragraph",
      "anchorText": "larger context containing the target",
      "boundaryStart": "start marker",
      "boundaryEnd": "end marker",
      "editType": "replace",
      "beforeText": "EXACT text from document to replace",
      "afterText": "new replacement text",
      "constraints": ["preserve structure"]
    }
  ]
}`;

  const response = await openai.chat.completions.create({
    model: 'gpt-4o',
    messages: [{ role: 'user', content: prompt }],
    response_format: { type: 'json_object' },
    temperature: 0.1,
  });

  const result = JSON.parse(response.choices[0].message.content || '{}');
  return result.editSpecs || [];
}

// ========================
// 3. Structural Executor (Code)
// ========================

function executeEdits(
  document: ParsedDocument,
  editSpecs: EditSpec[]
): { executions: EditExecution[]; modifiedXml: any } {
  const executions: EditExecution[] = [];
  const modifiedXml = JSON.parse(JSON.stringify(document.xml)); // Deep clone

  const docKey = findKey(modifiedXml, 'w:document', 'document')!;
  const bodyKey = findKey(modifiedXml[docKey], 'w:body', 'body')!;
  const body = modifiedXml[docKey][bodyKey];
  const pKey = findKey(body, 'w:p', 'p')!;
  const paragraphs = Array.isArray(body[pKey]) ? body[pKey] : [body[pKey]];

  for (const spec of editSpecs) {
    const execution: EditExecution = {
      editSpecId: spec.changeItemId,
      anchorFound: false,
      applied: false,
      actualBefore: '',
      actualAfter: '',
      locationIndex: -1,
    };

    // Find paragraph containing anchor_text
    let targetParagraphIdx = -1;
    let targetParagraph: any = null;

    for (let i = 0; i < paragraphs.length; i++) {
      const pText = extractText(paragraphs[i]);
      // Normalize for matching
      const normalizedPText = pText.replace(/\s+/g, ' ').trim();
      const normalizedAnchor = spec.anchorText.replace(/\s+/g, ' ').trim();

      if (normalizedPText.includes(normalizedAnchor) ||
          pText.includes(spec.anchorText) ||
          normalizedPText.includes(spec.beforeText.replace(/\s+/g, ' ').trim())) {
        targetParagraphIdx = i;
        targetParagraph = paragraphs[i];
        execution.anchorFound = true;
        execution.locationIndex = i;
        break;
      }
    }

    if (!execution.anchorFound) {
      execution.blockReason = `Anchor text not found: "${spec.anchorText.substring(0, 50)}..."`;
      executions.push(execution);
      continue;
    }

    // Extract current text
    execution.actualBefore = extractText(targetParagraph);

    // Apply edit based on type
    if (spec.editType === 'replace' || spec.editType === 'insert_before' || spec.editType === 'insert_after') {
      const success = applyTextReplacement(
        targetParagraph,
        spec.beforeText,
        spec.afterText,
        spec.editType
      );

      if (success) {
        execution.applied = true;
        execution.actualAfter = extractText(targetParagraph);
      } else {
        execution.blockReason = `Could not locate before_text: "${spec.beforeText.substring(0, 50)}..."`;
      }
    } else if (spec.editType === 'delete') {
      const success = applyTextReplacement(targetParagraph, spec.beforeText, '', 'replace');
      if (success) {
        execution.applied = true;
        execution.actualAfter = extractText(targetParagraph);
      } else {
        execution.blockReason = `Could not locate text to delete: "${spec.beforeText.substring(0, 50)}..."`;
      }
    }

    executions.push(execution);
  }

  return { executions, modifiedXml };
}

function applyTextReplacement(
  paragraph: any,
  beforeText: string,
  afterText: string,
  editType: string
): boolean {
  const rKey = findKey(paragraph, 'w:r', 'r');
  if (!rKey) return false;

  const runs = Array.isArray(paragraph[rKey]) ? paragraph[rKey] : [paragraph[rKey]];

  // Collect all text with run mapping
  let fullText = '';
  const runMap: Array<{ runIndex: number; startPos: number; endPos: number; run: any }> = [];

  for (let i = 0; i < runs.length; i++) {
    const run = runs[i];
    const text = extractText(run);
    runMap.push({
      runIndex: i,
      startPos: fullText.length,
      endPos: fullText.length + text.length,
      run,
    });
    fullText += text;
  }

  // Normalize for matching
  const normalizedFull = fullText.replace(/\s+/g, ' ');
  const normalizedBefore = beforeText.replace(/\s+/g, ' ');

  // Find position (try exact first, then normalized)
  let startPos = fullText.indexOf(beforeText);
  if (startPos === -1) {
    // Try normalized matching
    const normStart = normalizedFull.indexOf(normalizedBefore);
    if (normStart === -1) return false;
    startPos = normStart;
  }

  const endPos = startPos + beforeText.length;

  // Find which runs are affected
  const affectedRuns: number[] = [];
  for (const rm of runMap) {
    if (rm.endPos > startPos && rm.startPos < endPos) {
      affectedRuns.push(rm.runIndex);
    }
  }

  if (affectedRuns.length === 0) return false;

  // Simple case: all text in one run
  if (affectedRuns.length === 1) {
    const runIdx = affectedRuns[0];
    const run = runs[runIdx];
    const tKey = findKey(run, 'w:t', 't');
    if (tKey) {
      let currentText = '';
      if (typeof run[tKey] === 'string') {
        currentText = run[tKey];
      } else if (run[tKey] && run[tKey]._) {
        currentText = run[tKey]._;
      }

      let newText: string;
      if (editType === 'insert_before') {
        newText = currentText.replace(beforeText, afterText + beforeText);
      } else if (editType === 'insert_after') {
        newText = currentText.replace(beforeText, beforeText + afterText);
      } else {
        newText = currentText.replace(beforeText, afterText);
      }

      // Apply highlight to the modified run
      applyHighlight(run);

      if (typeof run[tKey] === 'string') {
        run[tKey] = newText;
      } else if (run[tKey]) {
        run[tKey]._ = newText;
        run[tKey].$ = run[tKey].$ || {};
        run[tKey].$['xml:space'] = 'preserve';
      }

      return true;
    }
  }

  // Complex case: text spans multiple runs - merge and replace
  const firstRunIdx = affectedRuns[0];
  const firstRun = runs[firstRunIdx];

  // Reconstruct the text with replacement
  let combinedText = '';
  for (const idx of affectedRuns) {
    combinedText += extractText(runs[idx]);
  }

  let newText: string;
  if (editType === 'insert_before') {
    newText = combinedText.replace(beforeText, afterText + beforeText);
  } else if (editType === 'insert_after') {
    newText = combinedText.replace(beforeText, beforeText + afterText);
  } else {
    newText = combinedText.replace(beforeText, afterText);
  }

  // Put all text in first run
  const tKey = findKey(firstRun, 'w:t', 't') || 'w:t';
  firstRun[tKey] = { _: newText, $: { 'xml:space': 'preserve' } };
  applyHighlight(firstRun);

  // Clear other affected runs
  for (let i = 1; i < affectedRuns.length; i++) {
    const idx = affectedRuns[i];
    const tk = findKey(runs[idx], 'w:t', 't');
    if (tk) {
      if (typeof runs[idx][tk] === 'string') {
        runs[idx][tk] = '';
      } else if (runs[idx][tk]) {
        runs[idx][tk]._ = '';
      }
    }
  }

  return true;
}

function applyHighlight(run: any): void {
  const rPrKey = findKey(run, 'w:rPr', 'rPr') || 'w:rPr';
  if (!run[rPrKey]) {
    run[rPrKey] = {};
  }
  run[rPrKey]['w:highlight'] = { $: { 'w:val': 'yellow' } };
}

// ========================
// 4. Verification Engine (Code)
// ========================

async function verifyChanges(
  originalDocument: ParsedDocument,
  modifiedBuffer: Buffer,
  editSpecs: EditSpec[],
  executions: EditExecution[]
): Promise<VerificationResult> {
  const specResults: SpecVerification[] = [];
  let allPassed = true;

  // Parse modified document
  let modifiedDoc: ParsedDocument;
  try {
    modifiedDoc = await parseDocument(modifiedBuffer);
  } catch (error) {
    return {
      status: 'FAIL',
      specResults: [],
      documentIntegrity: {
        valid: false,
        canOpen: false,
        structurePreserved: false,
      },
      summary: `Document integrity failed: ${error}`,
    };
  }

  const modifiedFullText = modifiedDoc.fullText;

  for (let i = 0; i < editSpecs.length; i++) {
    const spec = editSpecs[i];
    const execution = executions.find(e => e.editSpecId === spec.changeItemId);

    const verification: SpecVerification = {
      editSpecId: spec.changeItemId,
      afterTextFound: false,
      locationCorrect: false,
      boundaryRespected: true,
      evidence: {
        expectedText: spec.afterText,
        foundText: '',
        matchPercentage: 0,
      },
      status: 'FAIL',
    };

    if (!execution || !execution.applied) {
      verification.failureReason = execution?.blockReason || 'Edit was not applied';
      allPassed = false;
      specResults.push(verification);
      continue;
    }

    // Check if after_text exists in modified document
    const normalizedModified = modifiedFullText.replace(/\s+/g, ' ').toLowerCase();
    const normalizedAfter = spec.afterText.replace(/\s+/g, ' ').toLowerCase();

    if (normalizedModified.includes(normalizedAfter) || modifiedFullText.includes(spec.afterText)) {
      verification.afterTextFound = true;
      verification.evidence.foundText = spec.afterText;
      verification.evidence.matchPercentage = 100;
    } else {
      // Partial match check
      const words = spec.afterText.split(/\s+/).filter(w => w.length > 2);
      const foundWords = words.filter(w => normalizedModified.includes(w.toLowerCase()));
      verification.evidence.matchPercentage = Math.round((foundWords.length / words.length) * 100);
      verification.evidence.foundText = `Partial: ${foundWords.join(', ')}`;
    }

    // Check location correctness
    if (execution.locationIndex >= 0 && execution.locationIndex < modifiedDoc.paragraphs.length) {
      const modifiedParagraphText = modifiedDoc.paragraphs[execution.locationIndex].text;
      if (modifiedParagraphText.includes(spec.afterText) ||
          modifiedParagraphText.replace(/\s+/g, ' ').includes(spec.afterText.replace(/\s+/g, ' '))) {
        verification.locationCorrect = true;
      }
    }

    // Determine pass/fail
    if (verification.afterTextFound && verification.locationCorrect) {
      verification.status = 'PASS';
    } else {
      verification.status = 'FAIL';
      verification.failureReason = !verification.afterTextFound
        ? `after_text not found in document`
        : `after_text found but not at expected location`;
      allPassed = false;
    }

    specResults.push(verification);
  }

  // Check document structure preservation
  const structurePreserved = originalDocument.paragraphs.length === modifiedDoc.paragraphs.length;

  return {
    status: allPassed ? 'PASS' : 'FAIL',
    specResults,
    documentIntegrity: {
      valid: true,
      canOpen: true,
      structurePreserved,
    },
    summary: allPassed
      ? `All ${specResults.length} edit(s) verified successfully`
      : `${specResults.filter(s => s.status === 'PASS').length}/${specResults.length} edits passed verification`,
  };
}

// ========================
// Main Pipeline
// ========================

export async function editContract(
  documentBuffer: Buffer,
  userRequest: string
): Promise<ContractEditResponse> {
  const timing = {
    interpretMs: 0,
    specGenMs: 0,
    executeMs: 0,
    verifyMs: 0,
    totalMs: 0,
  };

  const startTotal = Date.now();

  // Parse document
  const document = await parseDocument(documentBuffer);

  // Step 1: Interpret changes
  const startInterpret = Date.now();
  const changeItems = await interpretChanges(document, userRequest);
  timing.interpretMs = Date.now() - startInterpret;

  // Check for blocked items - only block if NO candidate has sufficient confidence
  // Having an ambiguity note is NOT enough to block if a high-confidence candidate exists
  const hasValidCandidate = changeItems.some(ci =>
    ci.candidates.some(c => c.confidence >= 0.7)
  );

  if (!hasValidCandidate && changeItems.length > 0) {
    return {
      userRequest,
      changeItems,
      editSpecs: [],
      executions: [],
      verification: {
        status: 'BLOCKED',
        specResults: [],
        documentIntegrity: { valid: true, canOpen: true, structurePreserved: true },
        summary: 'All changes blocked - no candidate location found with confidence >= 70%',
      },
      timing: { ...timing, totalMs: Date.now() - startTotal },
    };
  }

  // Filter to only change items with valid candidates
  const validChangeItems = changeItems.filter(ci =>
    ci.candidates.some(c => c.confidence >= 0.7)
  );

  // Step 2: Generate edit specs (only for valid change items)
  const startSpec = Date.now();
  const editSpecs = await generateEditSpecs(document, validChangeItems, userRequest);
  timing.specGenMs = Date.now() - startSpec;

  // Step 3: Execute edits
  const startExecute = Date.now();
  const { executions, modifiedXml } = executeEdits(document, editSpecs);
  timing.executeMs = Date.now() - startExecute;

  // Rebuild document
  const builder = new Builder();
  const modifiedXmlStr = builder.buildObject(modifiedXml);
  document.zip.file('word/document.xml', modifiedXmlStr);
  const modifiedBuffer = await document.zip.generateAsync({ type: 'nodebuffer' });

  // Step 4: Verify changes
  const startVerify = Date.now();
  const verification = await verifyChanges(document, modifiedBuffer, editSpecs, executions);
  timing.verifyMs = Date.now() - startVerify;

  timing.totalMs = Date.now() - startTotal;

  return {
    userRequest,
    changeItems,
    editSpecs,
    executions,
    verification,
    modifiedDocument: verification.status === 'PASS' ? modifiedBuffer : undefined,
    timing,
  };
}

export { parseDocument, ParsedDocument, ParsedParagraph };
