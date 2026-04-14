/* Auto-generated typed decision resolvers */
import type { StepId, DecisionStepId, FlowStep } from "./types";
import { decisionMap } from "./decisionMap";

/**
 * Single decide() router that delegates to the decision map.
 * Use this as the `decide` option in useFlowEngine().
 */
export const decisionExecuter = (step: FlowStep, statusValue: string): StepId => {
  const rule = decisionMap.steps[step.id as DecisionStepId];

  if (!rule) {
    throw new Error(`No decision rule found for step ${step.id}`);
  }

  const candidate = rule.Yes.find((t) => t.statusText === statusValue);

  if (candidate) {
    return candidate.id;
  }

  return rule.No;
};