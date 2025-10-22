// PeoplePicker.tsx
// React + TypeScript + Fluent UI v8 TagPicker + v9 Field
// Works in SPFx (uses SPHttpClient if provided) or standalone (fetch + request digest)

import * as React from "react";
import { Field } from "@fluentui/react-components";          // v9 for label/validation
import { TagPicker, ITag, IBasePickerSuggestionsProps } from "@fluentui/react"; // v8
// If you're not already bringing in v8 styles elsewhere, ensure fabric core or theme is loaded.

export type PrincipalType = 0 | 1 | 2 | 4 | 8 | 15; // None | User | DL | SecGroup | SPGroup | All

type PickerEntity = {
  Key: string;               // login/email/UPN/ID depending on directory
  DisplayText: string;       // primary display
  Description?: string;      // usually email
  EntityType?: string;
  IsResolved?: boolean;
  EntityData?: {
    Email?: string;
    MobilePhone?: string;
    Title?: string;          // job title
    Department?: string;
    PrincipalType?: string;
    AccountName?: string;
  };
};

export interface PeoplePickerProps {
  id: string;
  displayName?: string;
  placeholder?: string;
  isRequired?: boolean;
  disabled?: boolean;
  /** Single or multi-select */
  multi?: boolean;

  /** SharePoint site absolute URL, e.g. https://contoso.sharepoint.com/sites/HR */
  webUrl: string;

  /**
   * Limit which principal types are returned.
   * 1=User, 2=DL, 4=SecurityGroup, 8=SPGroup, 15=All. Default: 1 (User)
   */
  principalType?: PrincipalType;

  /** Maximum suggestions to request from the API */
  maxSuggestions?: number;

  /** Initial selected people (by email/login or DisplayText). You can pass either. */
  starterValue?: Array<{ key: string; text: string }>;

  /** Called with full resolved entities (raw API results) whenever selection changes */
  onChange?: (entities: PickerEntity[]) => void;

  /** (SPFx) Pass SPHttpClient to use built-in digest/headers handling. Optional. */
  spHttpClient?: any; // SPHttpClient
  /** (SPFx) Config enum, typically SPHttpClient.configurations.v1. Optional. */
  spHttpClientConfig?: any;

  /** When true, allow entering emails not found (we still try resolve). Default: false */
  allowFreeText?: boolean;
}

// ---- minimal in-memory digest cache (non-SPFx path) -------------------------
type DigestCache = {
  value: string;
  expiresAt: number; // epoch ms
};
const digestCache: Record<string, DigestCache> = {};

async function getRequestDigest(webUrl: string): Promise<string> {
  const now = Date.now();
  const cached = digestCache[webUrl];
  if (cached && cached.expiresAt > now + 5000) return cached.value;

  const resp = await fetch(`${webUrl}/_api/contextinfo`, {
    method: "POST",
    headers: {
      Accept: "application/json;odata=verbose",
    },
    body: "",
    credentials: "same-origin",
  });
  if (!resp.ok) throw new Error(`contextinfo failed: ${resp.status}`);
  const json = await resp.json();
  const digest = json?.d?.GetContextWebInformation?.FormDigestValue;
  const timeoutSec = json?.d?.GetContextWebInformation?.FormDigestTimeoutSeconds ?? 1800;
  digestCache[webUrl] = {
    value: digest,
    expiresAt: now + timeoutSec * 1000,
  };
  return digest;
}

// ---- call ClientPeoplePickerWebServiceInterface -----------------------------
async function searchPeopleViaREST(
  webUrl: string,
  query: string,
  principalType: PrincipalType,
  maxSuggestions: number,
  spHttpClient?: any,
  spHttpClientConfig?: any
): Promise<PickerEntity[]> {
  if (!query?.trim()) return [];

  // The API expects a JSON string for "queryParams"
  const pplPayload = {
    __metadata: { type: "SP.UI.ApplicationPages.ClientPeoplePickerQueryParameters" },
    QueryString: query,
    PrincipalSource: 15,                 // All sources
    PrincipalType: principalType ?? 1,   // Default Users only
    AllowMultipleEntities: true,
    MaximumEntitySuggestions: maxSuggestions || 25,
    SharePointGroupID: 0,
    Required: false,
  };

  const body = JSON.stringify({
    queryParams: JSON.stringify(pplPayload),
  });

  const url = `${webUrl}/_api/SP.UI.ApplicationPages.ClientPeoplePickerWebServiceInterface.clientPeoplePickerSearchUser`;

  // SPFx path: SPHttpClient manages digest & auth cookies
  if (spHttpClient && spHttpClientConfig) {
    const response = await spHttpClient.post(url, spHttpClientConfig, {
      headers: {
        Accept: "application/json;odata=verbose",
        "Content-Type": "application/json;odata=verbose",
        "odata-version": "3.0",
      },
      body,
    });
    if (!response.ok) throw new Error(`PeoplePicker search failed: ${response.status}`);
    const data = await response.json();
    const resultsStr: string = data?.d?.ClientPeoplePickerSearchUserResult ?? "[]";
    return JSON.parse(resultsStr) as PickerEntity[];
  }

  // Non-SPFx path: use fetch + request digest
  const digest = await getRequestDigest(webUrl);
  const resp = await fetch(url, {
    method: "POST",
    headers: {
      Accept: "application/json;odata=verbose",
      "Content-Type": "application/json;odata=verbose",
      "X-RequestDigest": digest,
      "odata-version": "3.0",
    },
    body,
    credentials: "same-origin",
  });
  if (!resp.ok) throw new Error(`PeoplePicker search failed: ${resp.status}`);
  const json = await resp.json();
  const resultsStr: string = json?.d?.ClientPeoplePickerSearchUserResult ?? "[]";
  return JSON.parse(resultsStr) as PickerEntity[];
}

// ---- tiny debounce hook -----------------------------------------------------
function useDebouncedCallback<T extends (...args: any[]) => any>(fn: T, delay = 250) {
  const ref = React.useRef<number>();
  return React.useCallback(
    (...args: Parameters<T>) =>
      new Promise<ReturnType<T>>((resolve) => {
        if (ref.current) window.clearTimeout(ref.current);
        ref.current = window.setTimeout(async () => {
          const out = await fn(...args);
          resolve(out);
        }, delay);
      }),
    [fn, delay]
  );
}

// ---- map PickerEntity -> ITag -----------------------------------------------
function toTag(entity: PickerEntity): ITag {
  const key = entity.Key ?? entity.EntityData?.AccountName ?? entity.EntityData?.Email ?? entity.DisplayText;
  const text =
    entity.DisplayText ||
    entity.EntityData?.Email ||
    entity.EntityData?.AccountName ||
    entity.Key ||
    "Unknown";
  return { key: key ?? text, name: text };
}

const suggestionsProps: IBasePickerSuggestionsProps = {
  suggestionsHeaderText: "People",
  noResultsFoundText: "No results",
  resultsMaximumNumber: 10,
  mostRecentlyUsedHeaderText: "",
};

// ---- Component --------------------------------------------------------------
export const PeoplePicker: React.FC<PeoplePickerProps> = ({
  id,
  displayName,
  placeholder,
  isRequired,
  disabled,
  multi = true,
  webUrl,
  principalType = 1,
  maxSuggestions = 25,
  starterValue,
  onChange,
  spHttpClient,
  spHttpClientConfig,
  allowFreeText = false,
}) => {
  const [selectedTags, setSelectedTags] = React.useState<ITag[]>(
    (starterValue ?? []).map((v) => ({ key: v.key, name: v.text }))
  );
  const [lastResolved, setLastResolved] = React.useState<PickerEntity[]>([]);

  const doSearch = React.useCallback(
    async (q: string) => {
      const results = await searchPeopleViaREST(
        webUrl,
        q,
        principalType,
        maxSuggestions,
        spHttpClient,
        spHttpClientConfig
      );
      setLastResolved(results);
      return results.map(toTag);
    },
    [webUrl, principalType, maxSuggestions, spHttpClient, spHttpClientConfig]
  );

  const debouncedSearch = useDebouncedCallback(doSearch, 250);

  const handleChange = React.useCallback(
    (items?: ITag[]) => {
      setSelectedTags(items ?? []);
      if (!onChange) return;

      // return full entities for selected tags
      const selectedSet = new Set((items ?? []).map((t) => String(t.key).toLowerCase()));
      const matched: PickerEntity[] = [];

      // match from lastResolved first
      for (const e of lastResolved) {
        const k =
          (e.Key ?? e.EntityData?.AccountName ?? e.EntityData?.Email ?? e.DisplayText ?? "").toLowerCase();
        if (selectedSet.has(k)) matched.push(e);
      }

      // if free text allowed, synthesize entities for tags we didn’t resolve
      if (allowFreeText) {
        for (const t of items ?? []) {
          const lk = String(t.key).toLowerCase();
          if (!matched.find((m) => (m.Key ?? "").toLowerCase() === lk)) {
            matched.push({
              Key: String(t.key),
              DisplayText: t.name,
              IsResolved: false,
              EntityData: { Email: /@/.test(String(t.key)) ? String(t.key) : undefined },
            });
          }
        }
      }

      onChange(matched);
    },
    [onChange, lastResolved, allowFreeText]
  );

  const picker = (
    <TagPicker
      onResolveSuggestions={(filter, selectedItems) => debouncedSearch(filter || "")}
      getTextFromItem={(tag) => tag.name}
      selectedItems={selectedTags}
      onChange={handleChange}
      pickerSuggestionsProps={suggestionsProps}
      inputProps={{ placeholder: placeholder ?? "Search people…" }}
      itemLimit={multi ? undefined : 1}
      disabled={disabled}
    />
  );

  const requiredMsg =
    isRequired && (selectedTags?.length ?? 0) === 0 ? "This is a required field and cannot be blank!" : undefined;

  return displayName ? (
    <Field label={displayName} validationMessage={requiredMsg} validationState={requiredMsg ? "error" : "none"}>
      {picker}
    </Field>
  ) : (
    picker
  );
};

// ---- Example usage ----------------------------------------------------------
/*
<PeoplePicker
  id="AssignedTo"
  displayName="Assign to"
  webUrl={this.props.context.pageContext.web.absoluteUrl}
  multi={true}
  principalType={1} // Users only
  starterValue={[{ key: "ada@example.com", text: "Ada Lovelace" }]}
  onChange={(entities) => {
    // For "people" fields in SharePoint list items, send numeric IDs if you have them,
    // otherwise set by "Claims" or "Key". For example, map to:
    // const valueForSharePoint = entities.map(e => ({ Key: e.Key }));
    console.log("Selected entities:", entities);
  }}
  // SPFx (optional, recommended inside SPFx):
  spHttpClient={this.props.context.spHttpClient}
  spHttpClientConfig={SPHttpClient.configurations.v1}
/>
*/

