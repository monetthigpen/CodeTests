// Auto-generated from Excel: Process Map tab

import type { ProcessMap } from './types';

export const processMap: ProcessMap = {
  startStepId: 'P100',
  steps: {
    P100: {
      id: 'P100',
      statusText: 'Submitted',
      shapeType: 'Start',
      function: 'EditForm',
      phase: 'Stage 1',
      edges: [{ to: 'P200' }]
    },
    P200: {
      id: 'P200',
      statusText: 'Approved',
      shapeType: 'End',
      function: 'EditForm',
      phase: 'Stage 2',
      edges: []
    },
    P300: {
      id: 'P300',
      statusText: 'Rejected',
      shapeType: 'End',
      function: 'EditForm',
      phase: 'Stage 2',
      edges: []
    }
  }
};