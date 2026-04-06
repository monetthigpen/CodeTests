/* Auto-generated from Excel: Process Map tab */

export type StepId =
  | 'P100'
  | 'P200'
  | 'P300'
  | 'P400'
  | 'P500'
  | 'P600'
  | 'P700';

export type DecisionStepId =
  | 'P200'
  | 'P300';

export type ShapeType = 'Start' | 'End' | 'Process' | 'Decision' | string;

/**
 * These are the business statuses your email/router logic cares about.
 * Keep this aligned with email.ts.
 */
export type StatusChoices =
  | 'Submitted'
  | 'Approved'
  | 'Rejected';

/**
 * Optional: if you still want the raw process-map text available separately,
 * this can help when reading the Excel-driven map.
 */
export type ProcessStatusText =
  | 'Open'
  | 'Decision'
  | 'Completed'
  | 'Rejected'
  | 'Invalid';

export interface FlowEdge {
  to: StepId;
  label?: string;
}

export interface FlowStep {
  id: StepId;
  statusText: string;
  shapeType: ShapeType;
  function?: string;
  phase?: string;
  edges: FlowEdge[];
}

export interface ProcessMap {
  steps: Record<StepId, FlowStep>;
  startStepId: StepId;
}

// export interface FlowContext { candidateStepId: StepId; }
// export type DecisionResult = StepId | string;
// export type DecisionResolver = (step: FlowStep, flowCtx: FlowContext, statusText: string) => DecisionResult;
// export type DecisionResolver = (step: FlowStep, statusText: string) => DecisionResult;

export interface DecisionYes {
  id: StepId;
  statusText: string;
}

export interface DecisionStep {
  decisionStepId: DecisionStepId;
  Yes: DecisionYes[];
  No: StepId;
}

export interface DecisionMap {
  steps: Record<DecisionStepId, DecisionStep>;
}

export type InternalFieldNames =
  | 'RequestTracker'
  | 'Internal_x0020_x0020_Comments'
  | 'Comment_x0020_History'
  | 'Status';

export type Requestor = {
  name: string;
  email: string;
  spId: number;
};

export type RequestHistoryEntry = {
  stepId: StepId;
  status: StatusChoices;
  timestamp: string;
  modifiedBy: string;
  ccOwnerName?: string;
  ccOwnerEmail?: string;
  requesterName?: string;
  requesterEmail?: string;
};

export type RequestTracker = {
  requestor: Requestor;
  history: RequestHistoryEntry[];
};