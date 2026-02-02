/**
 * Generalized Contract Editing Types
 *
 * This system is NOT a domain-specific tool (company/address/contact changer).
 * It is a general-purpose contract editing engine that:
 * - Understands document context
 * - Applies natural language change requests
 * - Preserves document structure
 * - Verifies changes accurately
 */

// ========================
// Change Item (from Change Interpreter)
// ========================

export type EditIntent = 'replace' | 'insert' | 'delete' | 'conditional';

export interface CandidateLocation {
  /** Which paragraph/cell/section contains the target */
  locationDescription: string;
  /** Exact text excerpt from document (for anchor matching) */
  excerptFromDocument: string;
  /** Why this location matches the user's request */
  rationale: string;
  /** Confidence score 0-1 */
  confidence: number;
}

export interface ChangeItem {
  /** Unique identifier for this change */
  id: string;
  /** Original user request fragment this addresses */
  userRequestFragment: string;
  /** What kind of edit is intended */
  intent: EditIntent;
  /** Up to 3 candidate locations in the document */
  candidates: CandidateLocation[];
  /** If ambiguous, explain why */
  ambiguityNote?: string;
}

// ========================
// Edit Spec (from Edit Spec Generator)
// ========================

export type TargetUnit = 'paragraph' | 'table_cell' | 'header' | 'footer' | 'run' | 'list_item';

export interface EditSpec {
  /** Reference to the ChangeItem this spec addresses */
  changeItemId: string;

  /** What structural unit contains the edit target */
  targetUnit: TargetUnit;

  /** Exact text to locate the edit position (from document) */
  anchorText: string;

  /** Where within the anchor to start the edit */
  boundaryStart: string;

  /** Where within the anchor to end the edit */
  boundaryEnd: string;

  /** Type of edit operation */
  editType: 'replace' | 'insert_before' | 'insert_after' | 'delete';

  /** The text to insert/replace with (empty for delete) */
  afterText: string;

  /** Text that existed before the edit (for verification) */
  beforeText: string;

  /** Structural elements that must be preserved */
  constraints: string[];
}

// ========================
// Execution Result
// ========================

export interface EditExecution {
  /** Which EditSpec was executed */
  editSpecId: string;

  /** Was the anchor found? */
  anchorFound: boolean;

  /** Was the edit applied? */
  applied: boolean;

  /** If not applied, why? */
  blockReason?: string;

  /** Actual text before edit (from document) */
  actualBefore: string;

  /** Actual text after edit (from document) */
  actualAfter: string;

  /** Paragraph/cell index where edit occurred */
  locationIndex: number;
}

// ========================
// Verification Result
// ========================

export interface VerificationResult {
  /** Overall pass/fail */
  status: 'PASS' | 'FAIL' | 'BLOCKED';

  /** Per-spec verification */
  specResults: SpecVerification[];

  /** Document integrity check */
  documentIntegrity: {
    valid: boolean;
    canOpen: boolean;
    structurePreserved: boolean;
  };

  /** Human-readable summary */
  summary: string;
}

export interface SpecVerification {
  editSpecId: string;

  /** Does after_text exist in the modified document? */
  afterTextFound: boolean;

  /** Was the change applied at the correct location? */
  locationCorrect: boolean;

  /** Were areas outside the boundary unchanged? */
  boundaryRespected: boolean;

  /** Specific evidence */
  evidence: {
    expectedText: string;
    foundText: string;
    matchPercentage: number;
  };

  /** PASS/FAIL for this spec */
  status: 'PASS' | 'FAIL';

  /** Reason for failure if applicable */
  failureReason?: string;
}

// ========================
// Complete Pipeline Response
// ========================

export interface ContractEditResponse {
  /** Original user request */
  userRequest: string;

  /** Decomposed change items */
  changeItems: ChangeItem[];

  /** Generated edit specifications */
  editSpecs: EditSpec[];

  /** Execution results */
  executions: EditExecution[];

  /** Verification results */
  verification: VerificationResult;

  /** Modified document buffer (only if PASS) */
  modifiedDocument?: Buffer;

  /** Pipeline timing */
  timing: {
    interpretMs: number;
    specGenMs: number;
    executeMs: number;
    verifyMs: number;
    totalMs: number;
  };
}
