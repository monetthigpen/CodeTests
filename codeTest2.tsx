import * as React from "react";
import { Field } from "@fluentui/react-components";
import { TagPicker, ITag, IBasePickerSuggestionsProps } from "@fluentui/react";

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
  isRequired?: boolean;
  submitting?: boolean;
  single?: boolean;
  starterValue?: { key: string; text: string } | { key: string; text: string }[];
  onChange?: (entities: PickerEntity[]) => void;
  spHttpClient?: any;
  spHttpClientConfig?: any;
}

const toTag = (e: PickerEntity): ITag => ({
  key: e.Key || e.EntityData?.AccountName || e.DisplayText,
  name: e.DisplayText || e.EntityData?.Email || e.Key,
});

const suggestionsProps: IBasePickerSuggestionsProps = {
  suggestionsHeaderText: "People",
  noResultsFoundText: "No results found",
  resultsMaximumNumber: 10,
};

export const PeoplePicker: React.FC<PeoplePickerProps> = (props) => {
  const {
    id: _id,
    displayName,
    className,
    description,
    placeholder,
    isRequired,
    submitting,
    single,
    starterValue,
    onChange,
    spHttpClient,
    spHttpClientConfig,
  } = props;

  // 🔹 HARD-CODE YOUR EXPLICIT URL HERE
  const webUrl = "https://yourtenant.sharepoint.com/sites/yoursite"; // <--- CHANGE THIS LINE
  const apiUrl = `${webUrl}/_api/SP.UI.ApplicationPages.ClientPeoplePickerWebServiceInterface.clientPeoplePickerSearchUser`;

  const starterArray = Array.isArray(starterValue)
    ? starterValue
    : starterValue
    ? [starterValue]
    : [];

  const [selectedTags, setSelectedTags] = React.useState<ITag[]>(
    starterArray.map((v) => ({ key: v.key, name: v.text }))
  );

  const searchPeople = async (query: string): Promise<ITag[]> => {
    if (!query.trim()) return [];

    const payload = {
      __metadata: { type: "SP.UI.ApplicationPages.ClientPeoplePickerQueryParameters" },
      QueryString: query,
      PrincipalSource: 15,
      PrincipalType: 15,
      AllowMultipleEntities: true,
      MaximumEntitySuggestions: 25,
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
          console.error("People Picker Error:", resp.status, resp.statusText, text);
          return [];
        }

        const data = await resp.json();
        const raw = data?.d?.ClientPeoplePickerSearchUserResult ?? "[]";
        const entities = JSON.parse(raw);
        return entities.map(toTag);
      }

      // fallback if spHttpClient isn't used
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
        console.error("Fetch Error:", resp.status, resp.statusText, text);
        return [];
      }

      const json = await resp.json();
      const raw = json?.d?.ClientPeoplePickerSearchUserResult ?? "[]";
      const entities = JSON.parse(raw);
      return entities.map(toTag);
    } catch (e) {
      console.error("People Picker Exception:", e);
      return [];
    }
  };

  const handleChange = (items?: ITag[]) => {
    setSelectedTags(items ?? []);
    if (onChange) {
      onChange(items as any);
    }
  };

  const requiredMsg =
    isRequired && selectedTags.length === 0 ? "This field is required." : undefined;

  return (
    <Field
      label={displayName}
      hint={description}
      validationMessage={requiredMsg}
      validationState={requiredMsg ? "error" : "none"}
    >
      <TagPicker
        className={className}
        disabled={submitting}
        itemLimit={single ? 1 : undefined}
        onResolveSuggestions={(filter) => searchPeople(filter || "")}
        getTextFromItem={(item) => item.name}
        selectedItems={selectedTags}
        onChange={handleChange}
        pickerSuggestionsProps={suggestionsProps}
        inputProps={{ placeholder: placeholder ?? "Search people…" }}
      />
    </Field>
  );
};

export default PeoplePicker;
