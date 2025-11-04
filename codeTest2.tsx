// PeoplePickerComponent.tsx
import * as React from "react";
import { Field } from "@fluentui/react-components";
import { TagPicker, ITag, IBasePickerSuggestionsProps } from "@fluentui/react";

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
  id?: string;
  displayName?: string;
  className?: string;
  description?: string;
  placeholder?: string;

  // canonical props
  isRequired?: boolean;
  submitting?: boolean;
  single?: boolean;
  disabled?: boolean;
  starterValue?: { key: string; text: string } | { key: string; text: string }[];
  onChange?: (entities: PickerEntity[]) => void;

  // SPFx (optional)
  spHttpClient?: any;
  spHttpClientConfig?: any;

  // tolerated extras from builder
  isrequired?: boolean;
  dateTimeFormat?: string;

  // optional tuning
  principalType?: PrincipalType;
  maxSuggestions?: number;
  allowFreeText?: boolean;
}

const toTag = (e: PickerEntity): ITag => ({
  key: e.Key || e.EntityData?.AccountName || e.EntityData?.Email || e.DisplayText,
  name: e.DisplayText || e.EntityData?.Email || e.Key,
});

const suggestionsProps: IBasePickerSuggestionsProps = {
  suggestionsHeaderText: "People",
  noResultsFoundText: "No results",
  resultsMaximumNumber: 10,
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
    single,
    disabled,
    starterValue,
    onChange,
    spHttpClient,
    spHttpClientConfig,
    principalType = 1,
    maxSuggestions = 25,
    allowFreeText = false,
    dateTimeFormat: _dateTimeFormat, // accepted but unused
  } = props;

  // Normalize required field
  const requiredEffective = (isRequired ?? isrequired) ?? false;

  // Explicit SharePoint Site URL
  const webUrl = "https://";
  const apiUrl = `${webUrl}/_api/SP.UI.ApplicationPages.ClientPeoplePickerWebServiceInterface.clientPeoplePickerSearchUser`;

  const starterArray = Array.isArray(starterValue)
    ? starterValue
    : starterValue
    ? [starterValue]
    : [];

  const [selectedTags, setSelectedTags] = React.useState<ITag[]>(
    starterArray.map((v) => ({ key: v.key, name: v.text }))
  );
  const [lastResolved, setLastResolved] = React.useState<PickerEntity[]>([]);

  const searchPeople = React.useCallback(
    async (query: string): Promise<ITag[]> => {
      if (!query.trim()) return [];

      const payload = {
        __metadata: { type: "SP.UI.ApplicationPages.ClientPeoplePickerQueryParameters" },
        QueryString: query,
        PrincipalSource: 15,
        PrincipalType: 15,
        AllowMultipleEntities: true,
        MaximumEntitySuggestions: maxSuggestions,
        SharePointGroupID: 0,
      };

      const body = JSON.stringify({ queryParams: JSON.stringify(payload) });

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

        // fallback (non-SPFx path)
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
    },
    [apiUrl, maxSuggestions, spHttpClient, spHttpClientConfig]
  );

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
    requiredEffective && selectedTags.length === 0
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
        searchPeople(filter || "").then((tags) =>
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

export default PeoplePicker;


