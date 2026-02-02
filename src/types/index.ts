/**
 * Core types for the Contract Editing System
 */

// ========================
// DOCUMENT CONTEXT TYPES
// ========================

export interface DocumentContext {
  documentType: string;
  definedTerms: DefinedTerm[];
  sectionStructure: SectionInfo[];
  totalParagraphs: number;
  hasHeaders: boolean;
  hasFooters: boolean;
  hasTables: boolean;
}

export interface DefinedTerm {
  term: string;
  definition: string;
  location: string;
}

export interface SectionInfo {
  sectionNumber: string;
  title: string;
  level: number;
  startParagraph: number;
  endParagraph: number;
}

// ========================
// CHANGE ITEM TYPES
// ========================

export type ChangeStatus =
  | 'PENDING'
  | 'HIGH_CONFIDENCE'
  | 'NEEDS_CONFIRMATION'
  | 'APPLIED'
  | 'FAILED'
  | 'BLOCKED';

export interface ChangeItem {
  changeId: string;
  userInstruction: string;
  description: string;
  status: ChangeStatus;
  candidateTargets: CandidateTarget[];
  selectedTarget?: CandidateTarget;
  editSpec?: EditSpec;
  failureReason?: string;
}

export interface CandidateTarget {
  targetId: string;
  anchorText: string;
  rationale: string;
  confidence: number; // 0-1
  location: TargetLocation;
}

export interface TargetLocation {
  paragraphIndex?: number;
  tableIndex?: number;
  rowIndex?: number;
  cellIndex?: number;
  headerFooterType?: 'header' | 'footer';
  sectionIndex?: number;
}

// ========================
// EDIT SPEC TYPES
// ========================

export type TargetUnit = 'paragraph' | 'table' | 'table_cell' | 'header' | 'footer' | 'run';

export type EditType = 'replace' | 'insert_before' | 'insert_after' | 'delete';

export interface EditSpec {
  changeId: string;
  targetUnit: TargetUnit;
  anchorText: string;
  boundaries: EditBoundaries;
  editType: EditType;
  beforeSnippet: string;
  afterText: string;
  highlight: boolean;
  constraints: EditConstraints;
}

export interface EditBoundaries {
  startOffset?: number;
  endOffset?: number;
  fullElement: boolean;
}

export interface EditConstraints {
  preserveNumbering: boolean;
  preserveDefinitions: boolean;
  preserveFormatting: boolean;
  preserveStyles: boolean;
}

// ========================
// DOCX ELEMENT TYPES
// ========================

export interface DocxParagraph {
  index: number;
  text: string;
  normalizedText: string;
  runs: DocxRun[];
  style?: string;
  numbering?: NumberingInfo;
  xmlElement: unknown;
}

export interface DocxRun {
  index: number;
  text: string;
  startOffset: number;
  endOffset: number;
  formatting: RunFormatting;
  xmlElement: unknown;
}

export interface RunFormatting {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  highlight?: string;
  fontSize?: number;
  fontName?: string;
}

export interface NumberingInfo {
  numId: string;
  level: number;
  format: string;
}

export interface DocxTable {
  index: number;
  rows: DocxTableRow[];
  xmlElement: unknown;
}

export interface DocxTableRow {
  index: number;
  cells: DocxTableCell[];
  xmlElement: unknown;
}

export interface DocxTableCell {
  rowIndex: number;
  cellIndex: number;
  paragraphs: DocxParagraph[];
  xmlElement: unknown;
}

export interface DocxHeaderFooter {
  type: 'header' | 'footer';
  sectionIndex: number;
  paragraphs: DocxParagraph[];
  xmlPath: string;
}

// ========================
// LOCATION MAPPING TYPES
// ========================

export interface LocatorResult {
  found: boolean;
  location?: ElementLocation;
  matchedText?: string;
  confidence: number;
  alternativeLocations?: ElementLocation[];
  error?: string;
}

export interface ElementLocation {
  type: 'paragraph' | 'run' | 'table_cell' | 'header' | 'footer';
  paragraphIndex?: number;
  runIndices?: number[];
  tableIndex?: number;
  rowIndex?: number;
  cellIndex?: number;
  headerFooterPath?: string;
  startCharOffset: number;
  endCharOffset: number;
}

// ========================
// EXECUTION RESULT TYPES
// ========================

export interface EditResult {
  changeId: string;
  success: boolean;
  beforeText: string;
  afterText: string;
  location: ElementLocation;
  error?: string;
}

// ========================
// VERIFICATION TYPES
// ========================

export type VerificationStatus = 'PASS' | 'FAIL';

export interface VerificationResult {
  status: VerificationStatus;
  checks: VerificationCheck[];
  overallError?: string;
}

export interface VerificationCheck {
  checkName: string;
  passed: boolean;
  details: string;
  evidence?: string;
}

export type FailureClassification =
  | 'TARGET_NOT_FOUND'
  | 'AMBIGUOUS_MATCH'
  | 'APPLY_FAILED'
  | 'VERIFICATION_FAILED'
  | 'DOCX_INTEGRITY_FAILED'
  | 'OBJECT_STRINGIFICATION'
  | 'TIMEOUT';

export interface QAEvidence {
  changeId: string;
  beforeSnippet: string;
  afterSnippet: string;
  verificationStatus: VerificationStatus;
  failureClassification?: FailureClassification;
  failureDetails?: string;
}

// ========================
// PIPELINE TYPES
// ========================

export interface PipelineRequest {
  documentBuffer: Buffer;
  userInstruction: string;
  fileName: string;
}

export interface PipelineResponse {
  success: boolean;
  documentBuffer?: Buffer;
  changeItems: ChangeItem[];
  qaEvidence: QAEvidence[];
  verificationResult: VerificationResult;
  blockedReason?: string;
  logs: PipelineLog[];
}

export interface PipelineLog {
  phase: string;
  timestamp: Date;
  message: string;
  data?: unknown;
}

// ========================
// API TYPES
// ========================

export interface EditRequest {
  instruction: string;
}

export interface EditResponse {
  success: boolean;
  downloadUrl?: string;
  changeItems: ChangeItemSummary[];
  qaEvidence: QAEvidence[];
  verificationResult: VerificationResult;
  blockedReason?: string;
}

export interface ChangeItemSummary {
  changeId: string;
  description: string;
  status: ChangeStatus;
  beforeSnippet?: string;
  afterSnippet?: string;
}

// ========================
// ERROR TYPES
// ========================

export class ContractEditError extends Error {
  constructor(
    message: string,
    public code: string,
    public classification?: FailureClassification,
    public recoverable: boolean = false
  ) {
    super(message);
    this.name = 'ContractEditError';
  }
}

export class BlockedError extends ContractEditError {
  constructor(message: string, classification: FailureClassification) {
    super(message, 'BLOCKED', classification, false);
    this.name = 'BlockedError';
  }
}
