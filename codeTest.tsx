export type StepId = 'P100' | 'P200' | 'P300';

export type StatusText =
  | 'Submitted'
  | 'Approved'
  | 'Rejected';

export type StatusChoice = StatusText;

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

export type DecisionStep = FlowStep;

export type DecisionMap = Record<StepId, DecisionStep>;

export type DecisionResolver = (
  currentStatus: StatusChoice,
  nextStatus?: StatusChoice
) => StepId;