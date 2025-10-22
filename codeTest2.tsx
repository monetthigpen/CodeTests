// PeoplePicker.tsx
// React + TypeScript + Fluent UI (v9 Field + v8 TagPicker)
// - Multi or single select
// - Starter values
// - Uses SP Client People Picker Web Service
// - SPFx or non-SPFx (auto-digest) compatible

import * as React from "react";
import { Field } from "@fluentui/react-components";
import { TagPicker, ITag, IBasePickerSuggestionsProps } from "@fluentui/react";

// Principal types per SP: 1=User, 2=DL, 4=SecGroup, 8=SPGroup, 15=All
export type PrincipalType = 0 | 1 | 2 | 4 | 8 | 15;

// Shape returned by ClientPeoplePickerWebServiceInterface
export type PickerEntity = {
  Key: string;               // UPN / email / claims key
  DisplayText: string;
  Description?: string;      // usually email
  EntityType?: string;
  IsResolved?: boolean;
  EntityData?: {
    Email?: string;
    MobilePhone?: string;
    Title?: string;
    Department?: string;
    PrincipalType?: string;
    AccountName?: string;    // i:0#.f|membership|user@contoso.com
  };
};

export interface PeoplePickerProps {
  id: string;
  displayName?: string;
  placeholder?: string;
  isRequired?: boolean;
  disabled?: boolean;

  /** Set false for single-select; true is default */
  multi?: boolean;

  /** Absolute site URL (e.g., https://contoso.sharepoint.com/sites/HR) */
  webUrl: string;

  /** 1=User (default), 15=All, etc. */
  principalType?: PrincipalType;

  /** Max number of suggestions the API should return */
  maxSuggestions?: number;

  /** Optional initial selection */
  starterValue?: Array<{ key: string; text: string }>;

  /** Emits raw People Picker entities whenever selection changes */
  onChange?: (entities: PickerEntity[]) => void;

  /** Optional SPFx client + config to use SPHttpClient (digest handled for you) */
  spHttpClient?: any;        // SPHttpClient
  spHttpClientConfig?: any;  // SPHttpClient.configurations.v1

  /** Allow free-text tags not resolved by SP (default false) */
  allowFreeText?: boolean;
}

/* ------------------------------------------------
   Minimal request-digest cache for non-SPFx usage
-------------------------------------------------*/
type DigestCache = { value: string; expiresAt: number };
const digestCache: Record<string, DigestCache> = {};

async function getRequestDigest(webUrl: string): Promise<string> {
  const now = Date.now();
  const cached = digestCache[webUrl];
  if (cached && cached.expiresAt > now + 5000) return cached.value;

  const resp = await fetch(`${webUrl}/_api/contextinfo`, {
    method: "POST",
    headers: { Accept: "application/json;odata=verbose" },
    body: "",
    credentials: "same-origin",
  });
  if (!resp.ok) throw new Error(`contextinfo failed: ${resp.status}`);
  const json = await resp.json();
  const digest = json?.d?.GetContextWebInformation?.FormDigestValue as string;
  const timeoutSec =
    (json?.d?.GetContextWebInformation?.FormDigestTimeoutSeconds as number) ?? 1800;

  digestCache[webUrl] = { value: digest, expiresAt: now + timeoutSec * 1000 };
  return digest;
}

/* ------------------------------------------------
   Call SP Client People Picker Web Service
-------------------------------------------------*/
async function searchPeopleViaREST(
  webUrl: string,
  query: string,
  principalType: PrincipalType,
  maxSuggestions: number,
  spHttpClient?: any,
  spHttpClientConfig?: any
): Promise<PickerEntity[]> {
  if (!query?.trim()) return [];

  // Payload must be JSON-stringified in "queryParams"
  const pplPayload = {
    __metadata: { type: "SP.UI.ApplicationPages.ClientPeoplePickerQueryParameters" },
    QueryString: query,
    PrincipalSource: 15,
    PrincipalType: principalType ?? 1,
    AllowMultipleEntities: true,
    MaximumEntitySuggestions: maxSuggestions || 25,
    SharePointGroupID: 0,
    Required: false,
  };

  const body = JSON.stringify({ queryParams: JSON.stringify(pplPayload) });
  const url = `${webUrl}/_api/SP.UI.ApplicationPages.ClientPeoplePickerWebServiceInterface.clientPeoplePickerSearchUser`;

  // SPFx path (SPHttpClient handles digest)
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

  // Non-SPFx path (manual digest)
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

/* ------------------------------------------------
   Debounce helper for async functions
   Returns Promise<TResult> (no nested Promise)
-------------------------------------------------*/
function useDebouncedAsync<TArgs extends any[], TResult>(
  fn: (...args: TArgs) => Promise<TResult>,
  delay = 250
) {
  const timer = React.useRef<number>();
  return React.useCallback(
    (...args: TArgs): Promise<TResult> =>
      new Promise((resolve) => {
        if (timer.current) window.clearTimeout(timer.current);
        timer.current = window.setTimeout(async () => {
          const result = await fn(...args);
          resolve(result);
        }, delay);
      }),
    [fn, delay]
  );
}

/* ------------------------------------------------
   Helpers
-------------------------------------------------*/
function toTag(entity: PickerEntity): ITag {
  const key =
    entity.Key ||
    entity.EntityData?.AccountName ||
    entity.EntityData?.Email ||
    entity.DisplayText;
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

/* ------------------------------------------------
   Component
-------------------------------------------------*/
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
    async (q: string): Promise<ITag[]> => {
      if (!q?.trim()) return [];
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

  const debouncedSearch = useDebouncedAsync(doSearch, 250);

  const handleChange = React.useCallback(
    (items?: ITag[]) => {
      const next = items ?? [];
      setSelectedTags(next);

      if (!onChange) return;

      // Build set of selected keys for quick match
      const selectedKeys = new Set(next.map((t) => String(t.key).toLowerCase()));
      const matched: PickerEntity[] = [];

      // Prefer entities from the last search batch
      for (const e of lastResolved) {
        const k =
          (e.Key ??
            e.EntityData?.AccountName ??
            e.EntityData?.Email ??
            e.DisplayText ??
            "").toLowerCase();
        if (selectedKeys.has(k)) matched.push(e);
      }

      // Synthesize for any free-text tags not resolved
      if (allowFreeText) {
        for (const t of next) {
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

  const requiredMsg =
    isRequired && (selectedTags?.length ?? 0) === 0
      ? "This is a required field and cannot be blank!"
      : undefined;

  return (
    <>
      {displayName ? (
        <Field
          label={displayName}
          validationMessage={requiredMsg}
          validationState={requiredMsg ? "error" : "none"}
        >
          <TagPicker
            onResolveSuggestions={(filter, selected) =>
              debouncedSearch(filter || "").then((tags) =>
                tags.filter(
                  (t) => !(selected ?? []).some((s) => String(s.key) === String(t.key))
                )
              )
            }
            getTextFromItem={(t) => t.name}
            selectedItems={selectedTags}
            onChange={handleChange}
            pickerSuggestionsProps={suggestionsProps}
            inputProps={{ placeholder: placeholder ?? "Search people…" }}
            itemLimit={multi ? undefined : 1}
            disabled={disabled}
          />
        </Field>
      ) : (
        <TagPicker
          onResolveSuggestions={(filter, selected) =>
            debouncedSearch(filter || "").then((tags) =>
              tags.filter(
                (t) => !(selected ?? []).some((s) => String(s.key) === String(t.key))
              )
            )
          }
          getTextFromItem={(t) => t.name}
          selectedItems={selectedTags}
          onChange={handleChange}
          pickerSuggestionsProps={suggestionsProps}
          inputProps={{ placeholder: placeholder ?? "Search people…" }}
          itemLimit={multi ? undefined : 1}
          disabled={disabled}
        />
      )}
    </>
  );
};

export default PeoplePicker;

/*
Usage (SPFx):

<PeoplePicker
  id="AssignedTo"
  displayName="Assign to"
  webUrl={this.props.context.pageContext.web.absoluteUrl}
  multi={true}                 // or false for single
  principalType={1}            // 1=Users only, 15=All
  starterValue={[{ key: "ada@example.com", text: "Ada Lovelace" }]}
  spHttpClient={this.props.context.spHttpClient}
  spHttpClientConfig={SPHttpClient.configurations.v1}
  onChange={(entities) => {
    // Map to what your list expects, e.g. claims Keys:
    // const value = entities.map(e => ({ Key: e.Key }));
    console.log("Selected:", entities);
  }}
/>

Usage (non-SPFx, modern/classic page):
<PeoplePicker id="Ppl" webUrl="https://contoso.sharepoint.com/sites/HR" />
*/

