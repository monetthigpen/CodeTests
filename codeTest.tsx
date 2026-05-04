export const StatusChoices = [
  'Submitted',
  'Approved',
  'Rejected'
] as const;

export type StatusChoice = typeof StatusChoices[number];

export type StatusText = StatusChoice;

export type StepId = 'P100' | 'P200' | 'P300';

export type DecisionStepId = StepId;

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

export type DecisionMap = {
  steps: Record<DecisionStepId, DecisionStep>;
};

export type DecisionResolver = (
  currentStatus: StatusChoice,
  nextStatus?: StatusChoice
) => DecisionStepId;