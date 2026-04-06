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
  decisionStepId: StepId;
  Yes: DecisionYes[];
  No: StepId;
}

export interface DecisionMap {
  steps: Record<DecisionStepId, DecisionStep>;
}


