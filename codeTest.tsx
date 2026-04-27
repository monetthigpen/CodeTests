import * as React from "react";
import { Button, Spinner } from "@fluentui/react-components";
import { DynamicFormContext } from "@spfx-monorepo/shared-library/dist/cjs/components/DynamicFormContext";
import { postSPRestAPI, ReturnDataProps } from "@spfx-monorepo/shared-library/dist/cjs/Utils/postSPRestAPI";
import { FormCustomizerContext } from "@microsoft/sp-listview-extensibility";
import type { UserProps } from "@spfx-monorepo/shared-library/dist/cjs/Utils/types";
import {
  buildEmail,
  EmailPayload,
  EmailRouterContext,
  FlowBody,
  FlowResult,
  sendEmail
} from "../flowscaffold/email";
import { evaluateFieldRules } from "@spfx-monorepo/shared-library/dist/cjs/Utils/formRulesEngine";

interface ButtonProps {
  OnSubmit: (data: boolean) => void;
  submitting: boolean;
  formContext: FormCustomizerContext;
  selectedType: string;
}

type PplPickerStorage = {
  email: string;
  fullName: string;
};

type StepResult = {
  ok: boolean;
  value?: EmailRouterContext;
};

export default function SaveComponent(props: ButtonProps): JSX.Element {
  const ctx = DynamicFormContext();

  const [isHidden, setIsHidden] = React.useState<boolean>(false);
  const [isDisabled, setIsDisabled] = React.useState<boolean>(false);
  const [spinnerHidden, setSpinnerHidden] = React.useState<boolean>(true);
  const [spinnerLabel, setSpinnerLabel] = React.useState<string>("");

  const btnId = "btnSubmit";

  const FORM_MODE_DISPLAY = 4;
  const FORM_MODE_EDIT = 6;
  const FORM_MODE_NEW = 8;

  const showSpinner = (label: string) => {
    setSpinnerLabel(label);
    setSpinnerHidden(false);
  };

  const hideSpinner = () => {
    setSpinnerLabel("");
    setSpinnerHidden(true);
  };

  const endSubmitUI = () => {
    props.OnSubmit(false);
    hideSpinner();
  };

  const buildItemEndpoint = (itemID: number): string => {
    return `${props.formContext.pageContext.site.serverRelativeUrl}/_api/web/lists/GetByTitle('${props.formContext.list.title}')/items(${itemID})`;
  };

  const getRequestConfig = () => {
    if (ctx.FormMode === FORM_MODE_NEW) {
      return {
        uri: `${props.formContext.pageContext.site.serverRelativeUrl}/_api/web/lists/GetByTitle('${props.formContext.list.title}')/items`,
        method: "POST" as const
      };
    }

    const id = props.formContext.item?.ID as number;

    return {
      uri: `${props.formContext.pageContext.site.serverRelativeUrl}/_api/web/lists/GetByTitle('${props.formContext.list.title}')/items(${id})`,
      method: "PATCH" as const
    };
  };

  const waitForPeoplePicker = async (): Promise<boolean> => {
    let counter = 0;
    const maxIterations = 16;
    const intervalMs = 1000;

    return new Promise((resolve) => {
      const intervalID = setInterval(() => {
        const pplType = { dir: "OUT" };
        const stillResolving = Boolean(ctx.GlobalPplPicker(pplType as any));

        counter++;

        if (stillResolving === false || counter >= maxIterations) {
          clearInterval(intervalID);

          if (stillResolving === true) {
            resolve(false);
            return;
          }

          resolve(true);
        }
      }, intervalMs);
    });
  };

  const validateForm = async (): Promise<boolean> => {
    const res: any = ctx.GlobalReturnData();

    const errorItems = res?.errorItems ?? {};
    const listData = res?.listData ?? {};
    const frmData = res?.frmData ?? {};
    const requiredElements: string[] = res?.requiredItems ?? [];

    for (const field of requiredElements) {
      const value = ctx.FormMode === FORM_MODE_EDIT ? frmData[field] : listData[field];

      if (value === null || value === undefined || value === "") {
        alert("Please fill in all required fields");
        return false;
      }
    }

    if (Object.keys(errorItems).length > 0) {
      alert("Please review highlighted fields and try again!");
      return false;
    }

    return true;
  };

  const uploadAttachments = async (attachments: any[], itemID: number) => {
    if (!attachments || attachments.length === 0) return [];

    const itemUrl = buildItemEndpoint(itemID);

    const headers: HeadersInit = {
      Accept: "application/json",
      "Content-Type": "application/octet-stream"
    };

    const uploads = attachments.map((att) => {
      const fileUrl = `${itemUrl}/AttachmentFiles/add(FileName='${att.name}')`;

      const apiCall = {
        uri: fileUrl,
        body: att.content,
        context: props.formContext
      };

      return postSPRestAPI(apiCall, headers);
    });

    return await Promise.allSettled(uploads);
  };

  const stripAttachments = (data: Record<string, any>) => {
    const copy = { ...data };
    delete copy.attachments;
    return copy;
  };

  const saveItem = async (): Promise<ReturnDataProps> => {
    const { uri, method } = getRequestConfig();

    const raw = ctx.listSubData ?? {};
    const deepCopy = JSON.parse(JSON.stringify(raw));
    const dataToSend = stripAttachments(deepCopy);

    const headers: HeadersInit =
      method === "PATCH"
        ? {
            Accept: "application/json;odata.metadata=full",
            "Content-Type": "application/json;odata.metadata=full",
            "X-HTTP-Method": "MERGE",
            "IF-MATCH": "*"
          }
        : {
            Accept: "application/json;odata.metadata=full",
            "Content-Type": "application/json;odata.metadata=full"
          };

    const apiCall = {
      uri,
      body: JSON.stringify(dataToSend),
      context: props.formContext
    };

    const result = await postSPRestAPI(apiCall, headers);

    if (result.status !== 201 && result.status !== 204) {
      throw result;
    }

    return result;
  };

  const getCostCenterOwner = (): PplPickerStorage | null => {
    const formData = ctx.FormData as Record<string, any>;

    const ownerObj =
      formData["Cost_x0020_Center_x0020_Owner"] ??
      formData["CostCenterOwner"] ??
      null;

    if (ownerObj?.email && ownerObj?.fullName) {
      return {
        email: ownerObj.email,
        fullName: ownerObj.fullName
      };
    }

    if (ownerObj?.secondaryText && ownerObj?.name) {
      return {
        email: ownerObj.secondaryText,
        fullName: ownerObj.name
      };
    }

    if (ownerObj?.text && ownerObj?.secondaryText) {
      return {
        email: ownerObj.secondaryText,
        fullName: ownerObj.text
      };
    }

    return null;
  };

  const getCurrentUser = (): UserProps => ctx.curUserInfo;

  const stepHandler = async (): Promise<StepResult> => {
    const currentUser = getCurrentUser();
    const ccOwner = getCostCenterOwner();

    if (ctx.FormMode === FORM_MODE_NEW) {
      const isCostCenterOwner =
        !!ccOwner?.email &&
        ccOwner.email.toLowerCase() === currentUser.userEmail.toLowerCase();

      ctx.listSubData = {
        ...ctx.listSubData,
        Title: currentUser.userFullName
      };

      if (isCostCenterOwner) {
        ctx.listSubData = {
          ...ctx.listSubData,
          Status: "Approved"
        };
      }

      const emailCtx: EmailRouterContext = {
        status: isCostCenterOwner ? "Approved" : "Submitted",
        requesterName: currentUser.userFullName,
        requesterEmail: currentUser.userEmail,
        requestTypeText: props.selectedType,
        itemID: props.formContext.item?.ID ?? 0,
        formContext: props.formContext,
        ccOwnerName: ccOwner?.fullName,
        ccOwnerEmail: ccOwner?.email,
        isCostCenterOwner
      };

      return { ok: true, value: emailCtx };
    }

    const status = ctx.FormData?.Status as string | undefined;

    if (status === "Approved") {
      ctx.listSubData = {
        ...ctx.listSubData,
        Status: "Approved"
      };

      const emailCtx: EmailRouterContext = {
        status: "Approved",
        requesterName: currentUser.userFullName,
        requesterEmail: currentUser.userEmail,
        requestTypeText: props.selectedType,
        itemID: props.formContext.item?.ID ?? 0,
        formContext: props.formContext
      };

      return { ok: true, value: emailCtx };
    }

    if (status === "Rejected") {
      ctx.listSubData = {
        ...ctx.listSubData,
        Status: "Rejected"
      };

      const emailCtx: EmailRouterContext = {
        status: "Rejected",
        requesterName: currentUser.userFullName,
        requesterEmail: currentUser.userEmail,
        requestTypeText: props.selectedType,
        itemID: props.formContext.item?.ID ?? 0,
        formContext: props.formContext
      };

      return { ok: true, value: emailCtx };
    }

    return { ok: true };
  };

  const handleSubmit = async (e: React.MouseEvent<HTMLButtonElement>) => {
    e.preventDefault();

    if (ctx.FormMode === FORM_MODE_DISPLAY) return;

    props.OnSubmit(true);
    showSpinner("Getting info...");

    try {
      const ppOk = await waitForPeoplePicker();

      if (!ppOk) {
        alert("Issue resolving PeoplePicker. Please try again");
        endSubmitUI();
        return;
      }

      showSpinner("Validating...");
      const valid = await validateForm();

      if (!valid) {
        endSubmitUI();
        return;
      }

      const stepHandlerOk = await stepHandler();

      if (!stepHandlerOk.ok) {
        endSubmitUI();
        return;
      }

      showSpinner("Saving...");
      const saveResult = await saveItem();

      let itemId = props.formContext.item?.ID as number;

      if (saveResult.status === 201) {
        const saveData = saveResult.data as any;
        itemId = saveData?.Id ?? saveData?.ID ?? itemId;
      }

      if (Object.prototype.hasOwnProperty.call(ctx.FormData ?? {}, "attachments")) {
        const sorted = await uploadAttachments(ctx.FormData.attachments, itemId);
        const anyFailed = sorted.some((x: any) => x.status === "rejected");

        if (anyFailed) {
          alert("Saved, however attachments had an issue.");
        }
      }

      if (stepHandlerOk.value) {
        const emailContext: EmailRouterContext = {
          ...stepHandlerOk.value,
          itemID: itemId
        };

        const build: EmailPayload[] | null = buildEmail(emailContext);

        if (build !== null) {
          const resultEmail: FlowResult<FlowBody> = await sendEmail(
            build,
            props.formContext
          );

          console.log(resultEmail);
        }
      }

      alert(
        saveResult.status === 204
          ? "Your changes are successfully submitted"
          : "Thank you for submitting Authorized Requestor"
      );

      endSubmitUI();
    } catch (error: any) {
      alert(error?.statusText ?? error);
      endSubmitUI();
    }
  };

  React.useEffect(() => {
    setIsDisabled(props.submitting);
  }, [props.submitting]);

  React.useEffect(() => {
    if (ctx.FormMode === FORM_MODE_DISPLAY) {
      setIsHidden(true);
    } else {
      const decision = evaluateFieldRules(btnId, {
        formMode: ctx.FormMode,
        formData: ctx.FormData,
        curUserInfo: ctx.curUserInfo,
        formConfigJson: ctx.formRules
      });

      if (decision.isDisabled !== undefined) {
        setIsDisabled(decision.isDisabled);
      }

      if (decision.isHidden !== undefined) {
        setIsHidden(decision.isHidden);
      } else {
        setIsHidden(false);
      }
    }
  }, [ctx]);

  return (
    <>
      <div
        className="fieldClass"
        style={{ display: isHidden ? "none" : "block", textAlign: "right" }}
      >
        <Button
          appearance="primary"
          id={btnId}
          title="Submit"
          onClick={handleSubmit}
          {...(isDisabled && { disabled: true })}
        >
          Submit
        </Button>
      </div>

      <div
        className="spinner-container"
        style={{ display: spinnerHidden ? "none" : "block" }}
      >
        <Spinner labelPosition="after" label={spinnerLabel} />
      </div>
    </>
  );
