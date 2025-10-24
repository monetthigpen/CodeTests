import * as React from "react";
import { Field } from "@fluentui/react-components";
import { TagPicker, ITag, IBasePickerSuggestionsProps } from "@fluentui/react";

// ----------------------------------------------------
// Types for SharePoint People Picker
// ----------------------------------------------------
export type PrincipalType = 0 | 1 | 2 | 4 | 8 | 15;

export interface PickerEntity {
  Key: string;
  DisplayText: string;
  Description?: string;
  EntityType?: string;
  IsResolved?: boolean;
  EntityData?: {
    Email?: string;
    Title?: string;
    Department?: string;
    AccountName?: string;
  };
}

export interface PeoplePickerProps {
  id: string;
  displayName?: string;
  className?: string;
  description?: string;
  placeholder?: string;

  isRequired?: boolean;
  submitting?: boolean;
  single?: boolean;
  disabled?: boolean;

  starterValue?: { key: string; text: string } | { key: string; text: string }[];

  onChange?: (entities: PickerEntity[]) => void;

  webUrl?: string;
  principalType?: PrincipalType;
  maxSuggestions?: number;
  spHttpClient?: any;
  spHttpClientConfig?: any;
  allowFreeText?: boolean;
}

// ----------------------------------------------------
// Request Digest (non-SPFx fallback)
// ----------------------------------------------------
type DigestCache = { value: string; expiresAt: number };
const digestCache: Record<string, DigestCache> = {};

async function getRequestDigest(webUrl: string): Promise<string> {
  const now = Date.now();
  const cached = digestCache[webUrl];
  if (cached && cached.expiresAt > now + 5000) return cached.value;

  const resp = await fetch(`${webUrl}/_api/contextinfo`, {
    method: "POST",
    headers: { Accept: "application/json;odata=verbose" },
    credentials: "same-origin",
  });
  const json = await resp.json();
  const digest = json?.d?.GetContextWebInformation?.FormDigestValue;
  const timeout = json?.d?.GetContextWebInformation?.FormDigestTimeoutSeconds ?? 1800;
  digestCache[webUrl] = { value: digest, expiresAt: now + timeout * 1000 };
  return digest;
}

// ----------------------------------------------------
// Call SharePoint ClientPeoplePickerWebServiceInterface
// ----------------------------------------------------
async function searchPeopleViaREST(
  webUrl: string,
  query: string,
  principalType: PrincipalType,
  maxSuggestions: number,
  spHttpClient?: any,
  spHttpClientConfig?: any
): Promise<PickerEntity[]> {
  if (!query?.trim()) return [];

  const payload = {
    __metadata: { type: "SP.UI.ApplicationPages.ClientPeoplePickerQueryParameters" },
    QueryString: query,
    PrincipalSource: 15,
    PrincipalType: principalType ?? 1,
    AllowMultipleEntities: true,
    MaximumEntitySuggestions: maxSuggestions || 25,
  };
  const body = JSON.stringify({ queryParams: JSON.stringify(payload) });
  const url = `${webUrl}/_api/SP.UI.ApplicationPages.ClientPeoplePickerWebServiceInterface.clientPeoplePickerSearchUser`;

  if (spHttpClient && spHttpClientConfig) {
    const resp = await spHttpClient.post(url, spHttpClientConfig, {
      headers: {
        Accept: "application/json;odata=verbose",
        "Content-Type": "application/json;odata=verbose",
      },
      body,
    });
    const data = await resp.json();
    return JSON.parse(data?.d?.ClientPeoplePickerSearchUserResult ?? "[]");
  }

  const digest = await getRequestDigest(webUrl);
  const resp = await fetch(url, {
    method: "POST",
    headers: {
      Accept: "application/json;odata=verbose",
      "Content-Type": "application/json;odata=verbose",
      "X-RequestDigest": digest,
    },
    body,
    credentials: "same-origin",
  });
  const json = await resp.json();
  return JSON.parse(json?.d?.ClientPeoplePickerSearchUserResult ?? "[]");
}

// ----------------------------------------------------
// Async debounce (no nested Promises)
// ----------------------------------------------------
function useDebouncedAsync<TArgs extends any[], TResult>(
  fn: (...args: TArgs) => Promise<TResult>,
  delay = 250
) {
  const timer = React.useRef<number>();
  return React.useCallback(
    (...args: TArgs): Promise<TResult> =>
      new Promise((resolve) => {
        if (timer.current) clearTimeout(timer.current);
        timer.current = window.setTimeout(async () => resolve(await fn(...args)), delay);
      }),
    [fn, delay]
  );
}

// ----------------------------------------------------
// Utility: map entity → ITag
// ----------------------------------------------------
const toTag = (e: PickerEntity): ITag => ({
  key: e.Key || e.EntityData?.AccountName || e.EntityData?.Email || e.DisplayText,
  name: e.DisplayText || e.EntityData?.Email || e.Key,
});

const suggestionsProps: IBasePickerSuggestionsProps = {
  suggestionsHeaderText: "People",
  noResultsFoundText: "No results",
  resultsMaximumNumber: 10,
};

// ----------------------------------------------------
// Component
// ----------------------------------------------------
const PeoplePickerInner: React.FC<PeoplePickerProps> = (props) => {
  const {
    displayName,
    className,
    description,
    placeholder,
    isRequired,
    submitting,
    single,
    disabled,
    starterValue,
    onChange,
    webUrl: webUrlProp,
    principalType = 1,
    maxSuggestions = 25,
    spHttpClient,
    spHttpClientConfig,
    allowFreeText = false,
  } = props;

  const webUrl =
    webUrlProp ||
    (typeof window !== "undefined" &&
      (window as any)._spPageContextInfo?.webAbsoluteUrl);

  const starterArray = Array.isArray(starterValue)
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
        const key =
          (e.Key ??
            e.EntityData?.AccountName ??
            e.EntityData?.Email ??
            e.DisplayText ??
            "").toLowerCase();
        if (selectedKeys.has(key)) matched.push(e);
      }

      if (allowFreeText) {
        for (const t of next) {
          const key = String(t.key).toLowerCase();
          if (!matched.find((m) => (m.Key ?? "").toLowerCase() === key)) {
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
    isRequired && selectedTags.length === 0
      ? "This field is required."
      : undefined;

  const isDisabled = Boolean(submitting || disabled);
  const itemLimit = single ? 1 : undefined;

  const picker = (
    <TagPicker
      className={className}
      disabled={isDisabled}
      itemLimit={itemLimit}
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

// ----------------------------------------------------
// Memoized Export
// ----------------------------------------------------
export const PeoplePicker = React.memo(PeoplePickerInner);
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

