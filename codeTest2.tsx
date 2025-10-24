import * as React from "react";
import { Field } from "@fluentui/react-components"; // v9
import { TagPicker, ITag, IBasePickerSuggestionsProps } from "@fluentui/react"; // v8

// If you WILL pass SPHttpClient from SPFx usage, import the type (optional):
// import { SPHttpClient } from "@microsoft/sp-http";

// -----------------------------
// Types from SP People Picker
// -----------------------------
export type PrincipalType = 0 | 1 | 2 | 4 | 8 | 15;

export type PickerEntity = {
  Key: string;               // Claims/UPN/email
  DisplayText: string;
  Description?: string;
  EntityType?: string;
  IsResolved?: boolean;
  EntityData?: {
    Email?: string;
    MobilePhone?: string;
    Title?: string;
    Department?: string;
    PrincipalType?: string;
    AccountName?: string;    // e.g., i:0#.f|membership|user@contoso.com
  };
};

type StarterItem = { key: string; text: string };

// -----------------------------
// Props (matches your call-site)
// -----------------------------
export interface PeoplePickerProps {
  id: string;
  displayName?: string;
  className?: string;
  description?: string;
  placeholder?: string;

  // your booleans
  isRequired?: boolean;
  submitting?: boolean;
  single?: boolean; // if true => single select; if false/undefined => multi

  // initial selection
  starterValue?: StarterItem | StarterItem[];

  // optional overrides / integration
  onChange?: (entities: PickerEntity[]) => void;

  // SharePoint setup
  /** If omitted, we try window._spPageContextInfo.webAbsoluteUrl */
  webUrl?: string;

  /** Narrow results: 1=User (default), 15=All */
  principalType?: PrincipalType;

  /** Suggestion count (default 25) */
  maxSuggestions?: number;

  /** (Optional SPFx) Pass SPHttpClient + config to skip manual digest fetch */
  spHttpClient?: any;        // SPHttpClient
  spHttpClientConfig?: any;  // SPHttpClient.configurations.v1

  /** Allow unresolved free-text entries (default false) */
  allowFreeText?: boolean;
}

// -----------------------------
// Minimal digest cache (non-SPFx)
// -----------------------------
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

// -----------------------------
// Call SP Client People Picker
// -----------------------------
async function searchPeopleViaREST(
  webUrl: string,
  query: string,
  principalType: PrincipalType,
  maxSuggestions: number,
  spHttpClient?: any,
  spHttpClientConfig?: any
): Promise<PickerEntity[]> {
  if (!query?.trim()) return [];

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

// -----------------------------
// Debounced async helper
// -----------------------------
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

// -----------------------------
// Small helpers
// -----------------------------
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

// -----------------------------
// Component
// -----------------------------
export const PeoplePicker: React.FC<PeoplePickerProps> = (props) => {
  const {
    id,
    displayName,
    className,
    description,
    placeholder,
    isRequired,
    submitting,
    single,
    starterValue,
    onChange,
    webUrl: webUrlProp,
    principalType = 1,
    maxSuggestions = 25,
    spHttpClient,
    spHttpClientConfig,
    allowFreeText = false,
  } = props;

  // Resolve webUrl fallback (works on classic/modern/SPFx pages)
  const webUrl =
    webUrlProp ||
    (typeof window !== "undefined" &&
      (window as any)._spPageContextInfo?.webAbsoluteUrl);

  if (!webUrl) {
    // Soft warning – still render, but searches will fail until webUrl is provided.
    // You can throw here if you prefer hard failure:
    // throw new Error("PeoplePicker requires webUrl or _spPageContextInfo.webAbsoluteUrl.");
  }

  // Normalize starterValue to array of ITag
  const starterArray: StarterItem[] = Array.isArray(starterValue)
    ? starterValue
    : starterValue
    ? [starterValue]
    : [];

  const [selectedTags, setSelectedTags] = React.useState<ITag[]>(
    starterArray.map((v) => ({ key: v.key, name: v.text }))
  );
  const [lastResolved, setLastResolved] = React.useState<PickerEntity[]>([]);

  const doSearch = React.useCallback(
    async (q: string): Promise<ITag[]> => {
      if (!q?.trim() || !webUrl) return [];
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

      const selectedKeys = new Set(next.map((t) => String(t.key).toLowerCase()));
      const matched: PickerEntity[] = [];

      for (const e of lastResolved) {
        const k =
          (e.Key ??
            e.EntityData?.AccountName ??
            e.EntityData?.Email ??
            e.DisplayText ??
            "").toLowerCase();
        if (selectedKeys.has(k)) matched.push(e);
      }

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

  const itemLimit = single ? 1 : undefined;
  const isDisabled = !!submitting || props["disabled"] === true;

  const picker = (
    <TagPicker
      className={className}
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
      itemLimit={itemLimit}
      disabled={isDisabled}
    />
  );

  return displayName ? (
    <Field
      label={displayName}
      hint={description}
      validationMessage={requiredMsg}
      validationState={requiredMsg ? "error" : "none"}
    >
      {picker}
    </Field>
  ) : (
    picker
  );
};

export default PeoplePicker;

<PeoplePicker
  id={listColumns[i].name}
  displayName={listColumns[i].displayName}
  starterValue={starterVal}
  isRequired={listColumns[i].required}
  submitting={isSubmitting}
  single={!listColumns[i].multi}
  placeholder={listColumns[i].description}
  description={listColumns[i].description}
  className="elementsWidth"
/>

