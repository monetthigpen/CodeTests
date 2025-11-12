// PeoplePickerComponent.tsx
import * as React from "react";
import { Field } from "@fluentui/react-components";          // v9
import { TagPicker, ITag, IBasePickerSuggestionsProps } from "@fluentui/react"; // v8

/* ---------------------------------- Types --------------------------------- */

export type PrincipalType = 0 | 1 | 2 | 4 | 8 | 15;

export interface PickerEntity {
  Key: string;
  DisplayText?: string;
  EntityType?: string;
  EntityType?: string;
  IsResolved?: boolean;
  EntityData?: {
    Email?: string;
    AccountName?: string;
    Title?: string;
    Department?: string;
  };
}

export interface PeoplePickerProps {
  id: string;
  displayName?: string;
  className?: string;
  description?: string;
  placeholder?: string;

  /** builder passes either of these; normalize below */
  isRequired?: boolean;
  isrequired?: boolean;

  submitting?: boolean;
  /** match TagPicker API: multiselect controls single vs multi */
  multiselect?: boolean;
  disabled?: boolean;

  /** starter can be single or array, keep shape compatible with TagPicker */
  starterValue?:
    | { key: string | number; text: string }
    | { key: string | number; text: string }[];

  /** notify parent with resolved SharePoint-style entities */
  onChange?: (entities: PickerEntity[]) => void;

  /** optional knobs (defaults supplied) */
  principalType?: PrincipalType;  // 1 = User only
  maxSuggestions?: number;        // default 5
  allowFreeText?: boolean;        // default false

  /** optional SPFx client + config for first-class POST */
  spHttpClient?: any;
  spHttpClientConfig?: any;
}

/* ------------------------- Helpers / shared pieces ------------------------ */

const suggestionsProps: IBasePickerSuggestionsProps = {
  suggestionsHeaderText: "People",
  noResultsFoundText: "No results",
  resultsMaximumNumber: 5,
};

/** Make an ITag from a SharePoint people entity — never return undefined keys. */
const toTag = (e: PickerEntity): ITag => {
  const rawKey =
    e.Key ??
    e.EntityData?.AccountName ??
    e.EntityData?.Email ??
    e.DisplayText ??
    "";

  const rawName =
    e.DisplayText ??
    e.EntityData?.Email ??
    e.Key ??
    "(unknown)";

  return {
    key: String(rawKey),
    name: String(rawName),
  };
};

/* -------------------------------- Component -------------------------------- */

const PeoplePicker: React.FC<PeoplePickerProps> = (props) => {
  const {
    id: _id,
    displayName,
    className,
    description,
    placeholder,
    isRequired,
    isrequired,
    submitting,
    multiselect,
    disabled,
    starterValue,
    onChange,

    principalType = 1,
    maxSuggestions = 5,
    allowFreeText = false,

    spHttpClient,
    spHttpClientConfig,
  } = props;

  const requiredEffective = (isRequired ?? isrequired) ?? false;
  const isMulti = multiselect === true;

  // ---- Explicit web URL (match what you’ve been using) ----
  const webUrl = "https://amerihealthcaritas.sharepoint.com/sites/eokm";
  const apiUrl = `${webUrl}/_api/SP.UI.ApplicationPages.ClientPeoplePickerWebServiceInterface.clientPeoplePickerSearchUser`;

  // ---- Normalize starter(s) into ITag[] ----
  const starterArray = Array.isArray(starterValue)
    ? starterValue
    : starterValue
    ? [starterValue]
    : [];

  const normalizedStarter: ITag[] = (isMulti ? starterArray : starterArray.slice(-1)).map(v => ({
    key: String(v.key),
    name: v.text,
  }));

  const [selectedTags, setSelectedTags] = React.useState<ITag[]>(normalizedStarter);
  const [lastResolved, setLastResolved] = React.useState<PickerEntity[]>([]);

  /* -------------------------- Search (REST people API) ------------------------- */
  const searchPeople = React.useCallback(async (query: string): Promise<ITag[]> => {
    if (!query.trim()) return [];

    const body = JSON.stringify({
      queryParams: {
        __metadata: { type: "SP.UI.ApplicationPages.ClientPeoplePickerQueryParameters" },
        QueryString: query,
        PrincipalSource: 15,      // All
        AllowMultipleEntities: true,
        MaximumEntitySuggestions: maxSuggestions,
        PrincipalType: principalType, // 1 = Users
        AllUrlZones: false,
        AllowEmailAddresses: true,
      },
    });

    try {
      // Prefer SPHttpClient if supplied
      if (spHttpClient && spHttpClientConfig) {
        const resp = await spHttpClient.post(apiUrl, spHttpClientConfig, {
          headers: {
            Accept: "application/json;odata=verbose",
            "Content-Type": "application/json;odata=verbose",
            "odata-version": "3.0",
          },
          body,
        });
        if (!resp.ok) {
          const txt = await resp.text().catch(() => "");
          console.error("PeoplePicker SPHttpClient error:", resp.status, resp.statusText, txt);
          return [];
        }
        const data = await resp.json();
        const raw = data?.d?.ClientPeoplePickerSearchUserResult ?? "[]";
        const entities: PickerEntity[] = JSON.parse(raw);
        setLastResolved(entities);
        return entities.map(toTag);
      }

      // Fallback: classic fetch with request digest
      const digest =
        (document.getElementById("__REQUESTDIGEST") as HTMLInputElement)?.value || "";

      const resp = await fetch(apiUrl, {
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

      if (!resp.ok) {
        const txt = await resp.text().catch(() => "");
        console.error("PeoplePicker fetch error:", resp.status, resp.statusText, txt);
        return [];
      }

      const json = await resp.json();
      const raw = json?.d?.ClientPeoplePickerSearchUserResult ?? "[]";
      const entities: PickerEntity[] = JSON.parse(raw);
      setLastResolved(entities);
      return entities.map(toTag);
    } catch (e) {
      console.error("PeoplePicker exception:", e);
      return [];
    }
  }, [apiUrl, maxSuggestions, principalType, spHttpClient, spHttpClientConfig]);

  /* ------------------------ Change mapping back to entities ------------------------ */
  const handleChange = React.useCallback(
    (items: ITag[] = []) => {
      setSelectedTags(items);

      if (!onChange) return;

      // Build a quick lookup from resolved entities
      const resolvedByKey = new Map(
        lastResolved.map(e => [String(e.Key ?? e.EntityData?.AccountName ?? e.EntityData?.Email ?? e.DisplayText ?? "").toLowerCase(), e])
      );

      const result: PickerEntity[] = [];

      for (const t of items) {
        const lk = String(t.key).toLowerCase();

        const hit =
          lastResolved.find(e =>
            (e.Key ?? "").toLowerCase() === lk ||
            (e.EntityData?.AccountName ?? "").toLowerCase() === lk ||
            (e.EntityData?.Email ?? "").toLowerCase() === lk
          ) || resolvedByKey.get(lk);

        if (hit) {
          result.push(hit);
        } else if (allowFreeText) {
          // synthesize a minimal entity from free text/key
          result.push({
            Key: String(t.key),
            DisplayText: t.name,
            IsResolved: false,
            EntityData: { Email: /@/.test(String(t.key)) ? String(t.key) : undefined },
          });
        }
      }

      onChange(result);
    },
    [onChange, lastResolved, allowFreeText]
  );

  /* ------------------------- Picker rendering & behavior ------------------------- */

  const requiredMsg =
    requiredEffective && selectedTags.length === 0 ? "This field is required." : undefined;

  const isDisabled = Boolean(submitting || disabled);
  const itemLimit = isMulti ? undefined : 1; // v8 TagPicker respects itemLimit

  return displayName ? (
    <Field
      label={displayName}
      hint={description}
      validationMessage={requiredMsg}
      validationState={requiredMsg ? "error" : "none"}
    >
      <TagPicker
        className={className}
        disabled={isDisabled}
        itemLimit={itemLimit}
        onResolveSuggestions={(filter, selected) => {
          // If single-select and something already selected, don't offer more
          if (!isMulti && (selected?.length ?? 0) >= 1) return [];
          return searchPeople(filter || "").then(tags =>
            // filter out any already selected tags by key
            tags.filter(t => !(selected ?? []).some(s => String(s.key) === String(t.key)))
          );
        }}
        getTextFromItem={(t) => t.name}
        selectedItems={selectedTags}
        onChange={handleChange}
        pickerSuggestionsProps={suggestionsProps}
        inputProps={{ placeholder: placeholder ?? "Search people…" }}
      />
    </Field>
  ) : (
    <TagPicker
      className={className}
      disabled={isDisabled}
      itemLimit={itemLimit}
      onResolveSuggestions={(filter, selected) => {
        if (!isMulti && (selected?.length ?? 0) >= 1) return [];
        return searchPeople(filter || "").then(tags =>
          tags.filter(t => !(selected ?? []).some(s => String(s.key) === String(t.key)))
        );
      }}
      getTextFromItem={(t) => t.name}
      selectedItems={selectedTags}
      onChange={handleChange}
      pickerSuggestionsProps={suggestionsProps}
      inputProps={{ placeholder: placeholder ?? "Search people…" }}
    />
  );
};

export default PeoplePicker;
