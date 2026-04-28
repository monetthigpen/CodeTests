import * as React from "react";
import { useEffect, useState } from "react";
import { Button, Spinner } from "@fluentui/react-components";

import { DynamicFormContext } from "@spfx-monorepo/shared-library/dist/cjs/components/DynamicFormContext";

import postSPRestAPI from "@spfx-monorepo/shared-library/dist/cjs/Utils/postSPRestAPI";
import type { ReturnDataProps } from "@spfx-monorepo/shared-library/dist/cjs/Utils/postSPRestAPI";

import { evaluateFieldRules } from "@spfx-monorepo/shared-library/dist/cjs/Utils/formRulesEngine";

import { FormCustomizerContext } from "@microsoft/sp-listview-extensibility";