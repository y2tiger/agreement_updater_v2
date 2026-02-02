/**
 * OpenAI Contract Architect Agent
 *
 * Uses OpenAI to READ the document (read-only) and produce Edit Specs.
 * This agent is a PLANNER, not an executor.
 *
 * Responsibilities:
 * - Understand document structure, definitions, references
 * - Convert user instructions into atomic Change Items
 * - Locate candidate target sections for each Change Item
 * - Produce strict Edit Specs
 * - Never edits DOCX directly
 */

import OpenAI from 'openai';
import { v4 as uuidv4 } from 'uuid';
import {
  DocumentContext,
  ChangeItem,
  CandidateTarget,
  EditSpec,
  DefinedTerm,
  SectionInfo,
  TargetLocation,
  ContractEditError,
} from '../types';
import { ParsedDocx, extractPlainText } from '../utils/docx-parser';

const MAX_CHANGE_ITEMS = 10;

interface ArchitectResult {
  documentContext: DocumentContext;
  changeItems: ChangeItem[];
  logs: string[];
}

/**
 * OpenAI Contract Architect Agent
 */
export class ContractArchitectAgent {
  private openai: OpenAI;
  private model: string;
  private logs: string[] = [];

  constructor(apiKey: string, model: string = 'gpt-4-turbo-preview') {
    this.openai = new OpenAI({ apiKey });
    this.model = model;
  }

  private log(message: string): void {
    const timestamp = new Date().toISOString();
    this.logs.push(`[${timestamp}] ${message}`);
  }

  /**
   * Main entry point: analyze document and create edit specs
   */
  async analyze(
    parsedDocx: ParsedDocx,
    userInstruction: string
  ): Promise<ArchitectResult> {
    this.logs = [];
    this.log('Starting Contract Architect analysis');

    // PHASE 0: Global Context Scan
    this.log('PHASE 0: Global Context Scan');
    const documentContext = await this.scanDocumentContext(parsedDocx);

    // PHASE 1: Change Decomposition
    this.log('PHASE 1: Change Decomposition');
    const changeItems = await this.decomposeChanges(
      parsedDocx,
      userInstruction,
      documentContext
    );

    // PHASE 2: Candidate Targeting
    this.log('PHASE 2: Candidate Targeting');
    await this.findCandidateTargets(parsedDocx, changeItems, documentContext);

    // PHASE 3: Edit Spec Lock
    this.log('PHASE 3: Edit Spec Lock');
    await this.lockEditSpecs(changeItems, documentContext);

    this.log('Contract Architect analysis complete');

    return {
      documentContext,
      changeItems,
      logs: this.logs,
    };
  }

  /**
   * PHASE 0: Scan document to build context
   */
  private async scanDocumentContext(parsedDocx: ParsedDocx): Promise<DocumentContext> {
    const plainText = extractPlainText(parsedDocx);

    const prompt = `You are analyzing a legal contract document. Provide a structured analysis.

DOCUMENT TEXT:
${plainText.substring(0, 15000)}${plainText.length > 15000 ? '\n[TRUNCATED]' : ''}

Analyze this document and respond with a JSON object containing:
{
  "documentType": "string - type of contract (e.g., 'Employment Agreement', 'NDA', 'Service Agreement')",
  "definedTerms": [
    {"term": "string", "definition": "brief definition", "location": "where defined"}
  ],
  "sectionStructure": [
    {"sectionNumber": "string", "title": "string", "level": number}
  ]
}

Focus on:
1. Identifying the document type
2. Finding defined terms (capitalized terms with specific meanings)
3. Mapping the section/clause structure

Respond ONLY with valid JSON, no other text.`;

    try {
      const response = await this.openai.chat.completions.create({
        model: this.model,
        messages: [{ role: 'user', content: prompt }],
        temperature: 0,
        response_format: { type: 'json_object' },
      });

      const content = response.choices[0]?.message?.content;
      if (!content) {
        throw new Error('Empty response from OpenAI');
      }

      const analysis = JSON.parse(content);

      const context: DocumentContext = {
        documentType: analysis.documentType || 'Unknown',
        definedTerms: (analysis.definedTerms || []).map((dt: Record<string, string>) => ({
          term: dt.term,
          definition: dt.definition,
          location: dt.location,
        })),
        sectionStructure: (analysis.sectionStructure || []).map((s: Record<string, unknown>, i: number) => ({
          sectionNumber: s.sectionNumber as string || '',
          title: s.title as string || '',
          level: s.level as number || 0,
          startParagraph: i,
          endParagraph: i + 1,
        })),
        totalParagraphs: parsedDocx.paragraphs.length,
        hasHeaders: parsedDocx.headersFooters.some(hf => hf.type === 'header'),
        hasFooters: parsedDocx.headersFooters.some(hf => hf.type === 'footer'),
        hasTables: parsedDocx.tables.length > 0,
      };

      this.log(`Document type: ${context.documentType}`);
      this.log(`Found ${context.definedTerms.length} defined terms`);
      this.log(`Found ${context.sectionStructure.length} sections`);

      return context;
    } catch (error) {
      this.log(`Error in context scan: ${error}`);
      // Return minimal context on error
      return {
        documentType: 'Unknown',
        definedTerms: [],
        sectionStructure: [],
        totalParagraphs: parsedDocx.paragraphs.length,
        hasHeaders: parsedDocx.headersFooters.some(hf => hf.type === 'header'),
        hasFooters: parsedDocx.headersFooters.some(hf => hf.type === 'footer'),
        hasTables: parsedDocx.tables.length > 0,
      };
    }
  }

  /**
   * PHASE 1: Decompose user instruction into atomic change items
   */
  private async decomposeChanges(
    parsedDocx: ParsedDocx,
    userInstruction: string,
    documentContext: DocumentContext
  ): Promise<ChangeItem[]> {
    const plainText = extractPlainText(parsedDocx);

    const prompt = `You are a legal document editor. Break down the user's edit instruction into atomic changes.

DOCUMENT TYPE: ${documentContext.documentType}

DEFINED TERMS:
${documentContext.definedTerms.map(dt => `- ${dt.term}: ${dt.definition}`).join('\n')}

DOCUMENT STRUCTURE:
${documentContext.sectionStructure.map(s => `${s.sectionNumber} ${s.title}`).join('\n')}

DOCUMENT TEXT:
${plainText.substring(0, 12000)}${plainText.length > 12000 ? '\n[TRUNCATED]' : ''}

USER INSTRUCTION:
${userInstruction}

Break this instruction into atomic Change Items. Each Change Item should be:
1. A single, specific edit operation
2. Independent of other changes (can be applied in any order)
3. Clearly describable

Respond with a JSON object:
{
  "changeItems": [
    {
      "description": "Brief description of what needs to change",
      "userInstruction": "The part of the user instruction this addresses"
    }
  ]
}

RULES:
- Maximum ${MAX_CHANGE_ITEMS} Change Items
- Each change must be specific and actionable
- Do not propose changes that weren't requested
- If the instruction is ambiguous, create fewer, more conservative changes

Respond ONLY with valid JSON.`;

    try {
      const response = await this.openai.chat.completions.create({
        model: this.model,
        messages: [{ role: 'user', content: prompt }],
        temperature: 0,
        response_format: { type: 'json_object' },
      });

      const content = response.choices[0]?.message?.content;
      if (!content) {
        throw new Error('Empty response from OpenAI');
      }

      const result = JSON.parse(content);
      const items = result.changeItems || [];

      const changeItems: ChangeItem[] = items.slice(0, MAX_CHANGE_ITEMS).map(
        (item: { description: string; userInstruction: string }) => ({
          changeId: uuidv4(),
          description: item.description,
          userInstruction: item.userInstruction,
          status: 'PENDING' as const,
          candidateTargets: [],
        })
      );

      this.log(`Decomposed into ${changeItems.length} change items`);

      return changeItems;
    } catch (error) {
      this.log(`Error in change decomposition: ${error}`);
      throw new ContractEditError(
        'Failed to decompose changes',
        'DECOMPOSITION_FAILED',
        undefined,
        false
      );
    }
  }

  /**
   * PHASE 2: Find candidate targets for each change item
   */
  private async findCandidateTargets(
    parsedDocx: ParsedDocx,
    changeItems: ChangeItem[],
    documentContext: DocumentContext
  ): Promise<void> {
    const plainText = extractPlainText(parsedDocx);

    for (const changeItem of changeItems) {
      const prompt = `You are locating the exact text to modify in a legal document.

CHANGE TO MAKE:
${changeItem.description}

USER INSTRUCTION:
${changeItem.userInstruction}

DOCUMENT STRUCTURE:
${documentContext.sectionStructure.map(s => `${s.sectionNumber} ${s.title}`).join('\n')}

DOCUMENT TEXT:
${plainText.substring(0, 12000)}${plainText.length > 12000 ? '\n[TRUNCATED]' : ''}

Find up to 3 candidate locations where this change should be applied.
For EACH candidate, provide:
1. The EXACT text that should be modified (copy verbatim from document)
2. Why this is the right location
3. Your confidence (0.0 to 1.0)

Respond with a JSON object:
{
  "candidates": [
    {
      "anchorText": "EXACT text from document that needs to change - must be verbatim quote",
      "rationale": "Why this is the correct location",
      "confidence": 0.95
    }
  ],
  "needsConfirmation": false,
  "confirmationReason": "Only if needsConfirmation is true"
}

RULES:
- anchorText MUST be an exact quote from the document
- If you can't find a confident match, set needsConfirmation: true
- Prefer longer anchor text for unique matching
- Do not guess or fabricate text

Respond ONLY with valid JSON.`;

      try {
        const response = await this.openai.chat.completions.create({
          model: this.model,
          messages: [{ role: 'user', content: prompt }],
          temperature: 0,
          response_format: { type: 'json_object' },
        });

        const content = response.choices[0]?.message?.content;
        if (!content) {
          changeItem.status = 'NEEDS_CONFIRMATION';
          changeItem.failureReason = 'Empty response from OpenAI';
          continue;
        }

        const result = JSON.parse(content);

        if (result.needsConfirmation) {
          changeItem.status = 'NEEDS_CONFIRMATION';
          changeItem.failureReason = result.confirmationReason || 'Ambiguous target';
          this.log(`Change ${changeItem.changeId} needs confirmation: ${changeItem.failureReason}`);
          continue;
        }

        const candidates = result.candidates || [];
        changeItem.candidateTargets = candidates.map(
          (c: { anchorText: string; rationale: string; confidence: number }, i: number) => ({
            targetId: `${changeItem.changeId}-target-${i}`,
            anchorText: c.anchorText,
            rationale: c.rationale,
            confidence: c.confidence || 0.5,
            location: {}, // Will be filled by Locator Agent
          })
        );

        // Select the highest confidence target
        if (changeItem.candidateTargets.length > 0) {
          const bestTarget = changeItem.candidateTargets.reduce((best, current) =>
            current.confidence > best.confidence ? current : best
          );

          if (bestTarget.confidence >= 0.7) {
            changeItem.selectedTarget = bestTarget;
            changeItem.status = 'HIGH_CONFIDENCE';
            this.log(`Change ${changeItem.changeId} has high confidence target (${bestTarget.confidence})`);
          } else {
            changeItem.status = 'NEEDS_CONFIRMATION';
            changeItem.failureReason = `Best match confidence too low (${bestTarget.confidence})`;
            this.log(`Change ${changeItem.changeId} needs confirmation: low confidence`);
          }
        } else {
          changeItem.status = 'NEEDS_CONFIRMATION';
          changeItem.failureReason = 'No candidate targets found';
          this.log(`Change ${changeItem.changeId} needs confirmation: no targets found`);
        }
      } catch (error) {
        changeItem.status = 'NEEDS_CONFIRMATION';
        changeItem.failureReason = `Error finding targets: ${error}`;
        this.log(`Error finding targets for change ${changeItem.changeId}: ${error}`);
      }
    }
  }

  /**
   * PHASE 3: Lock Edit Specs for high-confidence changes
   */
  private async lockEditSpecs(
    changeItems: ChangeItem[],
    documentContext: DocumentContext
  ): Promise<void> {
    for (const changeItem of changeItems) {
      if (changeItem.status !== 'HIGH_CONFIDENCE' || !changeItem.selectedTarget) {
        continue;
      }

      const prompt = `You are creating an Edit Specification for a legal document change.

CHANGE DESCRIPTION:
${changeItem.description}

TARGET TEXT (exact quote from document):
${changeItem.selectedTarget.anchorText}

Create an Edit Spec with the exact replacement text.

Respond with a JSON object:
{
  "editType": "replace" | "insert_before" | "insert_after" | "delete",
  "beforeSnippet": "The exact text being changed (same as anchor)",
  "afterText": "The new text to replace it with",
  "preserveNumbering": true/false,
  "preserveDefinitions": true/false,
  "fullElement": true/false - true if replacing entire paragraph
}

RULES:
- beforeSnippet MUST match the target anchor text
- afterText should make the minimum necessary change
- Preserve legal document formatting and style
- Do not add content that wasn't requested
- If deleting, set afterText to empty string

Respond ONLY with valid JSON.`;

      try {
        const response = await this.openai.chat.completions.create({
          model: this.model,
          messages: [{ role: 'user', content: prompt }],
          temperature: 0,
          response_format: { type: 'json_object' },
        });

        const content = response.choices[0]?.message?.content;
        if (!content) {
          changeItem.status = 'FAILED';
          changeItem.failureReason = 'Failed to generate Edit Spec';
          continue;
        }

        const spec = JSON.parse(content);

        changeItem.editSpec = {
          changeId: changeItem.changeId,
          targetUnit: 'paragraph', // Will be refined by Locator Agent
          anchorText: changeItem.selectedTarget.anchorText,
          boundaries: {
            fullElement: spec.fullElement || false,
          },
          editType: spec.editType || 'replace',
          beforeSnippet: spec.beforeSnippet || changeItem.selectedTarget.anchorText,
          afterText: spec.afterText || '',
          highlight: true,
          constraints: {
            preserveNumbering: spec.preserveNumbering ?? true,
            preserveDefinitions: spec.preserveDefinitions ?? true,
            preserveFormatting: true,
            preserveStyles: true,
          },
        };

        this.log(`Edit Spec locked for change ${changeItem.changeId}`);
      } catch (error) {
        changeItem.status = 'FAILED';
        changeItem.failureReason = `Error generating Edit Spec: ${error}`;
        this.log(`Error generating Edit Spec for change ${changeItem.changeId}: ${error}`);
      }
    }
  }
}
