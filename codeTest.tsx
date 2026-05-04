export type StepId = 'P100' | 'P200' | 'P300';

export type StatusText =
  | 'Submitted'
  | 'Approved'
  | 'Rejected';

export type ShapeType = 'Start' | 'Process' | 'Decision' | 'End';

export type FlowStep = {
  id: StepId;
  statusText: StatusText;
  shapeType: ShapeType;
  function: string;
  phase: string;
  edges: Array<{ to: StepId; label?: string }>;
};

export type ProcessMap = {
  startStepId: StepId;
  steps: Record<StepId, FlowStep>;
};