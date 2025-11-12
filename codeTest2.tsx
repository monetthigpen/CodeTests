// PeoplePickerComponent.tsx
import * as React from "react";
import { Field } from "@fluentui/react-components";
import { TagPicker, ITag, IBasePickerSuggestionsProps } from "@fluentui/react";

export type PrincipalType = 0 | 1 | 2 | 4 | 8 | 15;

export interface PickerEntity {
  Key: string;
  DisplayText?: string;
  EntityType?: string;
  EntityType2?: string;
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
  isrequired?: boolean;      // tolerated alias per your builder
  submitting?: boolean;
  /** 👇 match TagPicker: pass true to allow multiple selections */
  multiselect?: boolean;     // <— use this, not `single`
  disabled?: boolean;
  starterValue?: { key: string; text: string } | { key: string; text: string }[];
  onChange?: (entities: PickerEntity[]) => void;

  // REST bits (kept exactly as you had)
  spHttpClient?: any;
  spHttpClientConfig?: any;
  principalType?: PrincipalType;
  maxSuggestions?: number;
  allowFreeText?: boolean;

  [key: string]: any;
}

const toTag = (e: PickerEntity): ITag => ({
  key: e.Key || e.EntityData?.AccountName || e.EntityData?.Email || e.DisplayText,
  name: e.DisplayText || e.EntityData?.Email || e.Key,
});

const suggestionsProps: IBasePickerSuggestionsProps = {
  suggestionsHeaderText: "People",
  noResultsFoundText: "No results",
  resultsMaximumNumber: 5,
};

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
    multiselect,             // 👈 source of truth like TagPicker
    disabled,
    starterValue,
    onChange,
    spHttpClient,
    spHttpClientConfig,
    principalType = 1,
    maxSuggestions = 25,
    allowFreeText = false,
  } = props;

  // ---- required flag (unchanged) ----
  const requiredEffective = (isRequired ?? isrequired) ?? false;

  // ---- explicit SharePoint URL (unchanged from your code) ----
  const webUrl  = "https://amerihealthcaritas.sharepoint.com/sites/eokm";
  const apiUrl  = `${webUrl}/_api/SP.UI.ApplicationPages.ClientPeoplePickerWebServiceInterface.clientPeoplePickerSearchUser`;

  // ---- SINGLE vs MULTI: behave like TagPicker ----
  const isMulti = multiselect === true;     // 👈 same meaning as your TagPicker prop

  // ---- starter normalization (keep 1 when single) ----
  const starterArray =
    Array.isArray(starterValue) ? starterValue :
    starterValue ? [starterValue] : [];

  const normalizedStarter = isMulti ? starterArray : starterArray.slice(-1);

  const [selectedTags, setSelectedTags] = React.useState<ITag[]>(
    normalizedStarter.map(v => ({ key: v.key, name: v.text }))
  );

  const [lastResolved, setLastResolved] = React.useState<PickerEntity[]>([]);

  // ---- Search SharePoint people (kept as you had it) ----
  const searchPeople = React.useCallback(async (query: string): Promise<ITag[]> => {
    if (!query?.trim()) return [];

    const body = JSON.stringify({
      queryParams: {
        __metadata: { type: "SP.UI.ApplicationPages.ClientPeoplePickerQueryParameters" },
        AllowEmailAddresses: true,
        AllowMultipleEntities: false,
        AllUrlZones: false,
        MaximumEntitySuggestions: maxSuggestions,
        QueryString: query,
        PrincipalSource: 1,
        PrincipalType: principalType,
      }
    });

    try {
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
          const text = await resp.text().catch(() => "");
          console.error("PeoplePicker SPHttpClient error:", resp.status, resp.statusText, text);
          return [];
        }

        const data = await resp.json();
        const raw = data?.d?.ClientPeoplePickerSearchUserResult ?? "[]";
        const entities: PickerEntity[] = JSON.parse(raw);
        setLastResolved(entities);
        return entities.map(toTag);
      }

      // ---- fallback fetch (unchanged) ----
      const digest = (document.getElementById("__REQUESTDIGEST") as HTMLInputElement)?.value || "";
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
        const text = await resp.text().catch(() => "");
        console.error("PeoplePicker fetch error:", resp.status, resp.statusText, text);
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

  // ---- onChange: clamp to 1 when NOT multiselect (TagPicker pattern) ----
  const handleChange = React.useCallback((items?: ITag[]) => {
    let next = items ?? [];

    if (!isMulti && next.length > 1) {
      // keep the most recent one (same approach you used in v9 TagPicker)
      next = [next[next.length - 1]];
    }

    setSelectedTags(next);

    if (!onChange) return;

    // Map selected tags back to the last resolved entities by key/email/account
    const keys = new Set(next.map(t => String(t.key).toLowerCase()));
    const matched: PickerEntity[] = [];
    for (const e of lastResolved) {
      const k =
        (e.Key ??
         e.EntityData?.AccountName ??
         e.EntityData?.Email ??
         e.DisplayText ??
         "").toLowerCase();
      if (keys.has(k)) matched.push(e);
    }
    onChange(matched);
  }, [isMulti, lastResolved, onChange]);

  // ---- suggestions: once 1 is chosen in single mode, stop suggesting ----
  const onResolveSuggestions = React.useCallback((filter: string, selected?: ITag[]) => {
    const sel = selected ?? selectedTags;
    if (!isMulti && sel.length >= 1) return [];
    const term = (filter ?? "").toString();
    return searchPeople(term).then(tags =>
      tags.filter(t => !(sel ?? []).some(s => String(s.key) === String(t.key)))
    );
  }, [isMulti, selectedTags, searchPeople]);

  const requiredMsg =
    requiredEffective && selectedTags.length === 0 ? "This field is required." : undefined;

  const isDisabled = Boolean(submitting || disabled);

  const picker = (
    <TagPicker
      className={className}
      disabled={isDisabled}
      onResolveSuggestions={onResolveSuggestions}
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

export default PeoplePicker;
