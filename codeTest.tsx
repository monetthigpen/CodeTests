//** statusText must match SharePoint Status column choices exactly */
import type { DecisionMap } from "./types";

export const decisionMap: DecisionMap = {
  steps: {
    P200: {
      decisionStepId: "P200",
      Yes: [
        { id: "P400", statusText: "Approved" }
      ],
      No: "P300"
    },

    P300: {
      decisionStepId: "P300",
      Yes: [
        { id: "P400", statusText: "Approved" }
      ],
      No: {
        id: "P600",
        statusText: "Rejected"
      }
    }
  }
};