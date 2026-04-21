import * as React from "react";
import { Button, Spinner } from "@fluentui/react-components";

import { DynamicFormContext } from "@spfx-monorepo/shared-library/dist/cjs/components/DynamicFormContext";
import { postSPRestAPI, ReturnDataProps } from "@spfx-monorepo/shared-library/dist/cjs/utils/postSPRestAPI";
import { evaluateFieldRules } from "@spfx-monorepo/shared-library/dist/cjs/utils/formRulesEngine";

import { FormCustomizerContext } from "@microsoft/sp-listview-extensibility";

import type { PplPicker } from "@spfx-monorepo/shared-library/dist/cjs/utils/id/preutils/types";
import type { UserProps } from "@spfx-monorepo/shared-library/dist/cjs/utils/types";

// ✅ FIXED PATHS (THIS WAS YOUR ISSUE)
import type { StepId } from "../flowscaffold/types";
import type { RequestTracker } from "../flowscaffold/types";

import { processMap } from "../flowscaffold/processMap";
import { createFlowEngine } from "../flowscaffold/engine";
import { decisionExecuter } from "../flowscaffold/deciders";

import {
  buildEmail,
  EmailPayload,
  EmailRouterContext,
  FlowBody,
  FlowResult,
  sendEmail
} from "../flowscaffold/email";