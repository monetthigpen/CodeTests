import {
  SubmitToCcOwnerVars,
  SubmitToEokmVars,
  ApprovedToEokmVars,
  RejectedToEokmVars
} from "./emailTypes";
import { FormCustomizerContext } from "@microsoft/sp-listview-extensibility";
import {
  IHttpClientOptions,
  HttpClient,
  HttpClientResponse
} from "@microsoft/sp-http";

/**
 * Keep this union aligned with the statuses you actually want to email on.
 * Do NOT confuse this with SharePoint Status choices.
 */
export type EmailStatus = "Submitted" | "Approved" | "Rejected";

/** Minimal info needed to build emails for this flow */
export interface EmailRouterContext {
  status: EmailStatus;
  requestTypeText: string;
  requesterName: string;
  requesterEmail?: string;
  itemID: string | number;
  formContext: FormCustomizerContext;

  ccOwnerName?: string;
  ccOwnerEmail?: string;

  eokmMailBox?: string;

  /**
   * true = submitter is the cost center owner
   * false = submitter is NOT the cost center owner
   */
  isCostCenterOwner?: boolean;
}

/** Flow response shape */
export type FlowBody = {
  success?: boolean;
  message?: string;
};

export type FlowResult<TBody = any> = {
  ok: boolean;
  status: number;
  statusText: string;
  body?: TBody;
  raw?: string;
  error?: string;
};

/** What gets posted to your flow */
export interface EmailPayload {
  to: string[];
  subject: string;
  html: string;
}

/* =========================================================
   HTML HELPERS
========================================================= */

function escapeHtml(value: string): string {
  return String(value).replace(/[&<>"']/g, (ch) =>
    ({
      "&": "&amp;",
      "<": "&lt;",
      ">": "&gt;",
      '"': "&quot;",
      "'": "&#39;"
    } as Record<string, string>)[ch]
  );
}

function p(innerHtml: string): string {
  return `<p style="margin:8px 0;font:11pt 'Segoe UI',Tahoma,Arial,sans-serif;color:#212529;">${innerHtml}</p>`;
}

function pRedItalic(innerHtml: string): string {
  return `<p style="margin:8px 0;font:11pt 'Segoe UI',Tahoma,Arial,sans-serif;color:red;font-style:italic;">${innerHtml}</p>`;
}

function strong(innerText: string): string {
  return `<strong>${escapeHtml(innerText)}</strong>`;
}

function spanBold(innerHtml: string): string {
  return `<span style="font-weight:bold;">${innerHtml}</span>`;
}

function spanRed(innerHtml: string): string {
  return `<span style="color:red;">${innerHtml}</span>`;
}

function link(href: string, label: string): string {
  return `<a href="${escapeHtml(href)}" target="_blank" rel="noopener noreferrer">${escapeHtml(label)}</a>`;
}

function container(innerHtml: string): string {
  return `
<table role="presentation" cellspacing="0" cellpadding="0" border="0" width="100%" style="border-collapse:collapse;">
  <tr>
    <td style="padding:12px 0;">
      ${innerHtml}
    </td>
  </tr>
</table>`.trim();
}

/* =========================================================
   SUBJECT BUILDERS
========================================================= */

function subjectSubmitToCcOwner(v: SubmitToCcOwnerVars): string {
  return `Authorized Requestor Request #${v.requestId} submitted for your approval`;
}

function subjectSubmitToEokm(v: SubmitToEokmVars): string {
  return `Authorized Requestor Request #${v.requestId} submitted by ${v.requesterFullName}`;
}

function subjectApprovedToEokm(v: ApprovedToEokmVars): string {
  return `Authorized Requestor Request #${v.requestId} approved`;
}

function subjectRejectedToEokm(v: RejectedToEokmVars): string {
  return `Authorized Requestor Request #${v.requestId} rejected`;
}

/* =========================================================
   HTML BUILDERS
========================================================= */

function renderSubmitToCcOwner(v: SubmitToCcOwnerVars): string {
  const inner = [
    p(`Dear ${escapeHtml(v.ccOwnerName)},`),
    p(
      `An ${strong(v.requestTypeText)} Authorized Requestor request has been submitted by ${strong(
        v.requesterFullName
      )} and is awaiting your approval.`
    ),
    p(`Request ID: ${strong(v.requestId)}`),
    p(`Status: ${spanBold(v.busStatus)}`),
    p(`Please open the ${link(v.editPageUrl, "form")} to review and approve or reject this request.`),
    p(`Thank you,`),
    p(`${spanBold("Knowledge Services Team")}`),
    pRedItalic(`Please do not reply to this auto-generated email.`)
  ].join("\n");

  return container(inner);
}

function renderSubmitToEokm(v: SubmitToEokmVars): string {
  const inner = [
    p(`Dear Knowledge Services Team,`),
    p(
      `An ${strong(v.requestTypeText)} Authorized Requestor request has been submitted by ${strong(
        v.requesterFullName
      )}.`
    ),
    p(`Request ID: ${strong(v.requestId)}`),
    p(`Status: ${spanBold(v.busStatus)}`),
    p(`Please open the ${link(v.editPageUrl, "form")} to review the request details.`),
    p(`Thank you,`),
    p(`${spanBold("Knowledge Services Team")}`),
    pRedItalic(`Please do not reply to this auto-generated email.`)
  ].join("\n");

  return container(inner);
}

function renderApprovedToEokm(v: ApprovedToEokmVars): string {
  const inner = [
    p(`Dear Knowledge Services Team,`),
    p(
      `An ${strong(v.requestTypeText)} Authorized Requestor request submitted by ${strong(
        v.requesterFullName
      )} has been approved.`
    ),
    p(`Request ID: ${strong(v.requestId)}`),
    p(`Status: ${spanBold(v.busStatus)}`),
    p(`Please open the ${link(v.editPageUrl, "form")} to review the approved request.`),
    p(`Thank you,`),
    p(`${spanBold("Knowledge Services Team")}`),
    pRedItalic(`Please do not reply to this auto-generated email.`)
  ].join("\n");

  return container(inner);
}

function renderRejectedToEokm(v: RejectedToEokmVars): string {
  const inner = [
    p(`Dear Knowledge Services Team,`),
    p(
      `An ${strong(v.requestTypeText)} Authorized Requestor request submitted by ${strong(
        v.requesterFullName
      )} has been ${spanRed(escapeHtml(v.busStatus))}.`
    ),
    p(`Request ID: ${strong(v.requestId)}`),
    p(`Please open the ${link(v.editPageUrl, "form")} to review the rejected request.`),
    p(`Thank you,`),
    p(`${spanBold("Knowledge Services Team")}`),
    pRedItalic(`Please do not reply to this auto-generated email.`)
  ].join("\n");

  return container(inner);
}

/* =========================================================
   ROUTER
========================================================= */

export function buildEmail(ctx: EmailRouterContext): EmailPayload[] | null {
  const editUrl = `https://${window.location.hostname}${ctx.formContext.list.serverRelativeUrl}/EditForm.aspx?ID=${ctx.itemID}`;

  const emails: EmailPayload[] = [];
  const eokmMailBox =
    ctx.eokmMailBox?.trim() || "EnterpriseOperationsKnowledgeManagement@amerihealthcaritas.com";

  switch (ctx.status) {
    case "Submitted": {
      const eokmVars: SubmitToEokmVars = {
        requesterFullName: ctx.requesterName,
        requestTypeText: ctx.requestTypeText,
        requestId: String(ctx.itemID),
        editPageUrl: editUrl,
        busStatus: "Submitted"
      };

      emails.push({
        to: [eokmMailBox],
        subject: subjectSubmitToEokm(eokmVars),
        html: renderSubmitToEokm(eokmVars)
      });

      if (!ctx.isCostCenterOwner) {
        if (!ctx.ccOwnerName || !ctx.ccOwnerEmail) {
          return null;
        }

        const ccOwnerVars: SubmitToCcOwnerVars = {
          ccOwnerName: ctx.ccOwnerName,
          ccOwnerEmail: ctx.ccOwnerEmail,
          requesterFullName: ctx.requesterName,
          requestTypeText: ctx.requestTypeText,
          requestId: String(ctx.itemID),
          editPageUrl: editUrl,
          busStatus: "Submitted"
        };

        emails.push({
          to: [ctx.ccOwnerEmail],
          subject: subjectSubmitToCcOwner(ccOwnerVars),
          html: renderSubmitToCcOwner(ccOwnerVars)
        });
      }

      return emails;
    }

    case "Approved": {
      const approvedVars: ApprovedToEokmVars = {
        requesterFullName: ctx.requesterName,
        requestTypeText: ctx.requestTypeText,
        requestId: String(ctx.itemID),
        editPageUrl: editUrl,
        busStatus: "Approved"
      };

      emails.push({
        to: [eokmMailBox],
        subject: subjectApprovedToEokm(approvedVars),
        html: renderApprovedToEokm(approvedVars)
      });

      return emails;
    }

    case "Rejected": {
      const rejectedVars: RejectedToEokmVars = {
        requesterFullName: ctx.requesterName,
        requestTypeText: ctx.requestTypeText,
        requestId: String(ctx.itemID),
        editPageUrl: editUrl,
        busStatus: "Rejected"
      };

      emails.push({
        to: [eokmMailBox],
        subject: subjectRejectedToEokm(rejectedVars),
        html: renderRejectedToEokm(rejectedVars)
      });

      return emails;
    }

    default:
      return null;
  }
}

/* =========================================================
   FLOW SENDER
========================================================= */

export async function sendEmail(
  payload: EmailPayload[],
  ctx: FormCustomizerContext
): Promise<FlowResult<FlowBody>> {
  const flowUrl =
    "https://default04afd16a7254e2f9260fce3985944.dc.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/YOUR_FLOW_URL_HERE/triggers/manual/paths/invoke?api-version=2016-10-01";

  const requestHeaders = {
    accept: "application/json",
    "content-type": "application/json"
  };

  const httpClientOptions: IHttpClientOptions = {
    body: JSON.stringify(payload),
    headers: requestHeaders
  };

  try {
    const response: HttpClientResponse = await ctx.httpClient.post(
      flowUrl,
      HttpClient.configurations.v1,
      httpClientOptions
    );

    const raw = await response.text();
    const body = safeJsonParse<FlowBody>(raw);

    return {
      ok: response.ok,
      status: response.status,
      statusText: response.statusText,
      raw,
      body
    };
  } catch (e: any) {
    return {
      ok: false,
      status: 0,
      statusText: "Client/Network error",
      error: e?.message ?? String(e)
    };
  }
}

function safeJsonParse<T>(text: string): T | undefined {
  const t = (text ?? "").trim();
  if (!t) return undefined;
  if (!(t.startsWith("{") || t.startsWith("["))) return undefined;

  try {
    return JSON.parse(t) as T;
  } catch {
    return undefined;
  }
}