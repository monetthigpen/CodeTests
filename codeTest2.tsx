// PeoplePickerComponent.tsx
import * as React from "react";
import { Field } from "@fluentui/react-components";  // v9
import { TagPicker, ITag, IBasePickerSuggestionsProps } from "@fluentui/react"; // v8

/* ------------------------------ Types ------------------------------ */

export type PrincipalType = 0 | 1 | 2 | 4 | 8 | 15;

export interface PickerEntity {
  Key: string;
  DisplayText: string;
  EntityType?: string;
  IsResolved: boolean;
  EntityData2?: {
    Email?: string;
    AccountName?: string;
    Title?: string;
    Department2?: string;
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

  /** match TagPicker API; multiselect controls single vs multi */
  multiselect?: boolean;
  disabled?: boolean;

  /** starter can be single or array; keep shape compatible with TagPicker */
  starterValue?:
    | { key: string | number; text: string }
    | { key: string | number; text: string }[];

  /** notify parent with resolved SharePoint-style entities */
  onChange?: (entities: PickerEntity[]) => void;

  /** optional knobs (defaults supplied) */
  principalType?: PrincipalType; // 1 = User only
  maxSuggestions?: number; // default 5
  allowFreeText?: boolean; // default false

  /** optional SPFx client + config for first-class POST */
  spHttpClient?: any;
  spHttpClientConfig?: any;
}

/* -------------------------- Helpers / shared pieces -------------------------- */

const suggestionsProps: IBasePickerSuggestionsProps = {
  suggestionsHeaderText: "People",
  noResultsFoundText: "No results",
  resultsMaximumNumber: 5,
};

/** Make an ITag from a SharePoint people entity – never return undefined keys. */
const toTag = (e: PickerEntity): ITag => {
  const rawKey =
    e.Key ??
    e.EntityData2?.AccountName ??
    e.EntityData2?.Email ??
    e.DisplayText ??
    "";

  const rawName =
    e.DisplayText ??
    e.EntityData2?.Email ??
    e.Key ??
    "(unknown)";

  return {
    key: String(rawKey),
    name: String(rawName),
  };
};

/* ------------------------------ Component ------------------------------ */

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

  // Explicit SharePoint site URL + PeoplePicker API URL
  const webUrl =
    "https://amerihealthcaritas.sharepoint.com/sites/eokm";
  const apiUrl = `${webUrl}/_api/SP.UI.ApplicationPages.ClientPeoplePickerWebServiceInterface.ClientPeoplePickerSearchUser`;

  // Normalize starter(s) into ITag[]
  const starterArray =
    Array.isArray(starterValue)
      ? starterValue
      : starterValue
      ? [starterValue]
      : [];

  const normalizedStarter: ITag[] = (isMulti ? starterArray : starterArray.slice(-1)).map(
    (v) => ({
      key: String(v.key),
      name: v.text,
    })
  );

  const [selectedTags, setSelectedTags] =
    React.useState<ITag[]>(normalizedStarter);
  const [lastResolved, setLastResolved] =
    React.useState<PickerEntity[]>([]);

  /* -------- EDIT FORM: hydrate starterValue that contains SPUserID(s) -------- */

  React.useEffect(() => {
    // Only run if we have starter values but haven't resolved anything yet
    if (!normalizedStarter.length || lastResolved.length > 0) {
      return;
    }

    // Try to interpret starter keys as numeric SPUserID values
    const numericIds = normalizedStarter
      .map((t) => Number(t.key))
      .filter((id) => !Number.isNaN(id));

    if (!numericIds.length) {
      return;
    }

    const abort = new AbortController();

    (async () => {
      const hydrated: PickerEntity[] = [];

      for (const id of numericIds) {
        try {
          const resp = await fetch(
            `${webUrl}/_api/web/getUserById(${id})`,
            {
              method: "GET",
              headers: {
                Accept: "application/json;odata=verbose",
              },
              signal: abort.signal,
            }
          );

          if (!resp.ok) {
            console.warn(
              "PeoplePicker getUserById failed",
              id,
              resp.status,
              resp.statusText
            );
            continue;
          }

          const json: any = await resp.json();
          const u = json.d;

          hydrated.push({
            Key: String(u.Id),
            DisplayText: u.Title,
            IsResolved: true,
            EntityType: "User",
            EntityData2: {
              Email: u.Email,
              AccountName: u.LoginName,
              Title: u.Title,
              Department2: u.Department || "",
            },
          });
        } catch (err) {
          if (abort.signal.aborted) {
            return;
          }
          console.error("PeoplePicker getUserById error", err);
        }
      }

      if (!hydrated.length) {
        return;
      }

      // Store resolved entities & show them in the picker
      setLastResolved(hydrated);
      const tags = hydrated.map(toTag);
      setSelectedTags(tags);

      if (onChange) {
        onChange(hydrated);
      }
    })();

    return () => abort.abort();
    // normalizedStarter is stable for a given render; webUrl is constant.
  }, [normalizedStarter, lastResolved.length, onChange, webUrl]);

  /* -------------------------- Search (REST people API) -------------------------- */

  const searchPeople = React.useCallback(
    async (query: string): Promise<ITag[]> => {
      if (!query.trim()) return [];

      const body = JSON.stringify({
        queryParams: {
          __metadata: {
            type:
              "SP.UI.ApplicationPages.ClientPeoplePickerQueryParameters",
          },
          AllowEmailAddresses: true,
          AllowMultipleEntities: isMulti,
          AllUrlZones: false,
          MaximumEntitySuggestions: maxSuggestions,
          QueryString: query,
          PrincipalSource: 1,
          PrincipalType: principalType,
        },
      });

      try {
        // Prefer SPFx client if provided
        if (spHttpClient && spHttpClientConfig) {
          const resp = await spHttpClient.post(
            apiUrl,
            spHttpClientConfig,
            {
              headers: {
                Accept: "application/json;odata=verbose",
                "Content-Type": "application/json;odata=verbose",
                "odata-version": "3.0",
              },
              body,
            }
          );

          if (!resp.ok) {
            const txt = await resp.text().catch(() => "");
            console.error(
              "PeoplePicker SPHttpClient error:",
              resp.status,
              resp.statusText,
              txt
            );
            return [];
          }

          const data: any = await resp.json();
          const raw = data.d?.ClientPeoplePickerSearchUser ?? "[]";
          const entities: PickerEntity[] = JSON.parse(raw);
          setLastResolved(entities);
          return entities.map(toTag);
        }

        // Fallback: classic fetch with request digest
        const digest = (document.getElementById(
          "__REQUESTDIGEST"
        ) as HTMLInputElement | null)?.value;

        const resp = await fetch(apiUrl, {
          method: "POST",
          headers: {
            Accept: "application/json;odata=verbose",
            "Content-Type": "application/json;odata=verbose",
            "X-RequestDigest": digest || "",
            "odata-version": "3.0",
          },
          body,
          credentials: "same-origin",
        });

        if (!resp.ok) {
          const txt = await resp.text().catch(() => "");
          console.error(
            "PeoplePicker fetch error:",
            resp.status,
            resp.statusText,
            txt
          );
          return [];
        }

        const json: any = await resp.json();
        const raw =
          json.d?.ClientPeoplePickerSearchUser ?? "[]";
        const entities: PickerEntity[] = JSON.parse(raw);
        setLastResolved(entities);
        return entities.map(toTag);
      } catch (e) {
        console.error("PeoplePicker exception:", e);
        return [];
      }
    },
    [
      apiUrl,
      isMulti,
      maxSuggestions,
      principalType,
      spHttpClient,
      spHttpClientConfig,
    ]
  );

  /* ------------------- Change mapping back to entities ------------------- */

  const handleChange = React.useCallback(
    (items: ITag[] = []) => {
      setSelectedTags(items);
      if (!onChange) return;

      // Build a quick lookup from resolved entities
      const resolvedByKey = new Map<
        string,
        PickerEntity
      >(
        lastResolved.map((e) => [
          (
            e.EntityData2?.AccountName ??
            e.EntityData2?.Email ??
            e.DisplayText ??
            ""
          ).toLowerCase(),
          e,
        ])
      );

      const result: PickerEntity[] = [];

      for (const t of items) {
        const lk = String(t.key).toLowerCase();

        const hit =
          Array.from(lastResolved).find(
            (e) =>
              (e.Key ?? "").toLowerCase() === lk ||
              (e.EntityData2?.AccountName ?? "")
                .toLowerCase() === lk ||
              (e.EntityData2?.Email ?? "").toLowerCase() === lk ||
              (e.DisplayText ?? "").toLowerCase() === lk
          ) || resolvedByKey.get(lk);

        if (hit) {
          result.push(hit);
        } else if (allowFreeText) {
          // synthesize a minimal entity from free text/key
          result.push({
            Key: String(t.key),
            DisplayText: t.name,
            IsResolved: false,
            EntityType: "User",
            EntityData2: {
              Email: /@/.test(String(t.key))
                ? String(t.key)
                : undefined,
            },
          });
        }
      }

      onChange(result);
    },
    [onChange, lastResolved, allowFreeText]
  );

  /* ---------------- Picker rendering & behavior ---------------- */

  const requiredMsg =
    requiredEffective && selectedTags.length === 0
      ? "This field is required."
      : undefined;

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
          if (!isMulti && (selected ?? []).length >= 1) return [];
          return searchPeople(filter || "").then((tags) =>
            tags.filter(
              (t) =>
                !(selected ?? []).some(
                  (s) => String(s.key) === String(t.key)
                )
            )
          );
        }}
        getTextFromItem={(t) => t.name}
        selectedItems={selectedTags}
        onChange={handleChange}
        pickerSuggestionsProps={suggestionsProps}
        inputProps={{
          placeholder: placeholder ?? "Search people…",
        }}
      />
    </Field>
  ) : (
    <TagPicker
      className={className}
      disabled={isDisabled}
      itemLimit={itemLimit}
      onResolveSuggestions={(filter, selected) => {
        if (!isMulti && (selected ?? []).length >= 1) return [];
        return searchPeople(filter || "").then((tags) =>
          tags.filter(
            (t) =>
              !(selected ?? []).some(
                (s) => String(s.key) === String(t.key)
              )
          )
        );
      }}
      getTextFromItem={(t) => t.name}
      selectedItems={selectedTags}
      onChange={handleChange}
      pickerSuggestionsProps={suggestionsProps}
      inputProps={{
        placeholder: placeholder ?? "Search people…",
      }}
    />
  );
};

export default PeoplePicker;


