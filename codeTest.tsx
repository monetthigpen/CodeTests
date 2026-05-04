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

export type DecisionTarget = {
  id: StepId;
  statusText: StatusText;
};

export type DecisionStep = {
  decisionStepId: DecisionStepId;
  Yes: DecisionTarget[];
  No: StepId;
};

export type DecisionMap = {
  steps: Partial<Record<DecisionStepId, DecisionStep>>;
};

export type DecisionResolver = (
  step: FlowStep,
  statusValue: string
) => StepId;




import type { DecisionMap } from './types';

export const decisionMap: DecisionMap = {
  steps: {
    P100: {
      decisionStepId: 'P100',
      Yes: [
        {
          id: 'P200',
          statusText: 'Approved'
        },
        {
          id: 'P300',
          statusText: 'Rejected'
        }
      ],
      No: 'P100'
    }
  }
};




import type { StepId, DecisionStepId, FlowStep } from './types';
import { decisionMap } from './decisionMap';

export const decisionExecuter = (
  step: FlowStep,
  statusValue: string
): StepId => {
  const rule = decisionMap.steps[step.id as DecisionStepId];

  if (!rule) {
    throw new Error(`No decision rule found for step ${step.id}`);
  }

  const candidate = rule.Yes.find(
    (t) => t.statusText === statusValue
  );

  if (candidate) {
    return candidate.id;
  }

  return rule.No;
};




import type { DecisionResolver, FlowStep, ProcessMap, StatusChoice, StepId } from './types';

const next = (stepId: StepId, statusValue: StatusChoice, options: TransitionOptions = {}): StepId | null => {