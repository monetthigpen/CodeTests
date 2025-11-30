// PeoplePickerComponent.tsx
import * as React from "react";
import { Field } from "@fluentui/react-components"; // v9
import {
  TagPicker,
  ITag,
  IBasePickerSuggestionsProps,
} from "@fluentui/react"; // v8
import { DynamicFormContext } from "./DynamicFormContext";

/* ------------------------------------------------
 * Types
 * ------------------------------------------------ */

export type PrincipalType = 0 | 1 | 2 | 4 | 8 | 15;

export interface PickerEntity {
  Key: string;
  DisplayText?: string;
  EntityType?: string;
  IsResolved?: boolean;
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

  isRequired?: boolean;
  isrequired2?: boolean;
  submitting?: boolean;

  /* match TagPicker: pass true to allow multiple selections */
  multiselect?: boolean;
  disabled?: boolean;

  /** optional default(s); kept for compatibility but NOT used for Edit hydrations */
  starterValue?:
    | { key: string | number; text: string }
    | { key: string | number; text: string }[];

  /** notify parent with SharePoint-style entities */
  onChange?: (entities: PickerEntity[]) => void;

  /* optional knobs (defaults supplied) */
  principalType?: PrincipalType; // 1 = user only
  maxSuggestions?: number; // default 5
  allowFreeText?: boolean; // default false

  /* optional SPFx client + config for first-class POST */
  spHttpClient?: any;
  spHttpClientConfig?: any;

  /* tolerate unknown props from builder */
  [key: string]: any;
}

/* ------------------------------------------------
 * Helpers / shared pieces
 * ------------------------------------------------ */

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
    "(unknown)";

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

/* ------------------------------------------------
 * Component
 * ------------------------------------------------ */

const PeoplePicker: React.FC<PeoplePickerProps> = (props) => {
  const ctx = React.useContext(DynamicFormContext);

  const {
    id,
    displayName,
    className,
    description,
    placeholder,
    isRequired,
    isrequired2,
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

  const requiredEffective = (isRequired ?? isrequired2) ?? false;
  const isMulti = multiselect === true;

  // Explicit site URL (same pattern you used earlier)
  const webUrl = "https://amerihealthcaritas.sharepoint.com/sites/eokm";
  const apiUrl =
    `${webUrl}/_api/SP.UI.ApplicationPages.ClientPeoplePickerWebServiceInterface.clientPeoplePickerSearchUser`;

  // --- Normalize starterValue into ITag[] (only used for NEW forms if provided) ---
  const starterArray: { key: string | number; text: string }[] =
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

  const [selectedTags, setSelectedTags] = React.useState<ITag[]>(normalizedStarter);
  const [lastResolved, setLastResolved] = React.useState<PickerEntity[]>([]);

  // Hidden input only for GlobalRefs (so it matches TagPicker pattern)
  const elemRef = React.useRef<HTMLInputElement | null>(null);

  /* ------------------------------------------------
   * Search (REST people API)
   * ------------------------------------------------ */

  const searchPeople = React.useCallback(
    async (query: string): Promise<ITag[]> => {
      if (!query.trim()) return [];

      const body = JSON.stringify({
        queryParams: {
          __metadata: {
            type: "SP.UI.ApplicationPages.ClientPeoplePickerQueryParameters",
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
            console.error(
              "PeoplePicker spHttpClient error:",
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
    [apiUrl, isMulti, maxSuggestions, principalType, spHttpClient, spHttpClientConfig]
  );

  /* ------------------------------------------------
   * Change mapping back to entities
   * ------------------------------------------------ */

  const handleChange = React.useCallback(
    (items?: ITag[]) => {
      const next = items ?? [];
      setSelectedTags(next);

      if (!onChange) return;

      const result: PickerEntity[] = [];

      // Build quick lookup from last resolved entities
      const resolvedByKey = new Map<string, PickerEntity>();
      for (const e of lastResolved) {
        const rawKey =
          e.Key ??
          e.EntityData2?.AccountName ??
          e.EntityData2?.Email ??
          e.DisplayText ??
          "";
        resolvedByKey.set(String(rawKey).toLowerCase(), e);
      }

      for (const t of next) {
        const lk = String(t.key).toLowerCase();
        const match = resolvedByKey.get(lk);

        if (match) {
          result.push(match);
          continue;
        }

        // Synthesize minimal entity from free text (if allowed)
        if (allowFreeText) {
          result.push({
            Key: String(t.key),
            DisplayText: t.name,
            IsResolved: false,
            EntityType: "User",
            EntityData2: {
              Email: /@/.test(String(t.key)) ? String(t.key) : undefined,
            },
          });
        }
      }

      onChange(result);
    },
    [allowFreeText, lastResolved, onChange]
  );

  /* ------------------------------------------------
   * EDIT / VIEW FORMS: hydrate PeoplePicker from ctx.FormData (SPUserID)
   * ------------------------------------------------ */

  React.useEffect(() => {
    // Only run for EditForm(6) or ViewForm(4)
    if (!(ctx.FormMode === 4 || ctx.FormMode === 6)) {
      return;
    }

    const fieldInternalName = id;
    const formData = ctx.FormData as any | undefined;
    if (!formData) return;

    // Look at <InternalName>Id, then <InternalName>Id, then <InternalName>StringId
    let rawValue: any = formData[`${fieldInternalName}Id`];
    if (rawValue === undefined) {
      const idProp = `${fieldInternalName}Id`;
      const stringIdProp = `${fieldInternalName}StringId`;
      rawValue = formData[idProp] ?? formData[stringIdProp];
    }

    if (rawValue === null || rawValue === undefined) return;

    // --- normalize whatever SP stored into numeric SPUserID ---
    const collectIds = (value: any): number[] => {
      if (value === null || value === undefined) return [];

      if (Array.isArray(value)) {
        const ids: number[] = [];
        for (const v of value) {
          if (v && typeof v === "object" && "Id" in v) {
            ids.push(Number((v as any).Id));
          } else {
            ids.push(Number(v));
          }
        }
        return ids.filter((id) => !Number.isNaN(id));
      }

      const str = String(value);
      const parts = str.split(/[;,#]/);
      return parts
        .map((p) => Number(p.trim()))
        .filter((id) => !Number.isNaN(id));
    };

    const numericIds = collectIds(rawValue);
    if (!numericIds.length) return;

    const abort = new AbortController();

    (async () => {
      const hydrated: PickerEntity[] = [];

      for (const userId of numericIds) {
        try {
          const resp = await fetch(
            `${webUrl}/_api/web/getUserById(${userId})`,
            {
              method: "GET",
              headers: { Accept: "application/json;odata=verbose" },
              signal: abort.signal,
            }
          );

          if (!resp.ok) {
            console.warn(
              "PeoplePicker getUserById failed",
              userId,
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
          if (abort.signal.aborted) return;
          console.error("PeoplePicker getUserById error", err);
        }
      }

      if (!hydrated.length) return;

      setLastResolved(hydrated);

      const tags = hydrated.map(toTag);
      setSelectedTags(tags);

      if (onChange) {
        onChange(hydrated);
      }
    })();

    return () => abort.abort();
  }, [ctx.FormMode, ctx.FormData, id, onChange, webUrl]);

  /* ------------------------------------------------
   * NEW FORM: reset picker state + register GlobalRefs
   * ------------------------------------------------ */

  React.useEffect(() => {
    // always register element for this field, like TagPicker does
    ctx.GlobalRefs(
      elemRef.current !== null ? elemRef.current : undefined,
      id
    );

    if (ctx.FormMode === 8) {
      // New form = start clean
      setSelectedTags([]);
      setLastResolved([]);
    }
  }, [ctx.FormMode, ctx, id]);

  /* ------------------------------------------------
   * Rendering
   * ------------------------------------------------ */

  const requiredMsg =
    requiredEffective && selectedTags.length === 0
      ? "This field is required."
      : undefined;

  const isDisabled = Boolean(submitting || disabled);
  const itemLimit = isMulti ? undefined : 1; // v8 TagPicker respects itemLimit

  const pickerProps = {
    className,
    disabled: isDisabled,
    itemLimit,
    onResolveSuggestions: (filter: string, selected?: ITag[]) => {
      const already = selected ?? [];
      return searchPeople(filter || "").then((tags) =>
        tags.filter(
          (t) =>
            !already.some((s) => String(s.key) === String(t.key))
        )
      );
    },
    getTextFromItem: (t: ITag) => t.name,
    selectedItems: selectedTags,
    onChange: handleChange,
    pickerSuggestionsProps: suggestionsProps,
  };

  return displayName ? (
    <Field
      label={displayName}
      hint={description}
      validationMessage={requiredMsg}
      validationState={requiredMsg ? "error" : "none"}
    >
      {/* hidden input purely so GlobalRefs has a DOM element like TagPicker */}
      <input
        ref={elemRef}
        id={id}
        type="text"
        style={{
          position: "absolute",
          left: "-10000px",
          width: 0,
          height: 0,
          opacity: 0,
          pointerEvents: "none",
        }}
        readOnly
      />
      <TagPicker
        {...pickerProps}
        inputProps={{ placeholder: placeholder ?? "Search people..." }}
      />
    </Field>
  ) : (
    <>
      <input
        ref={elemRef}
        id={id}
        type="text"
        style={{
          position: "absolute",
          left: "-10000px",
          width: 0,
          height: 0,
          opacity: 0,
          pointerEvents: "none",
        }}
        readOnly
      />
      <TagPicker
        {...pickerProps}
        inputProps={{ placeholder: placeholder ?? "Search people..." }}
      />
    </>
  );
};

export default PeoplePicker;
