// PeoplePickerComponent.tsx
import * as React from "react";
import { Field } from "@fluentui/react-components";          // v9
import {
  TagPicker,
  ITag,
  IBasePickerSuggestionsProps,
} from "@fluentui/react";                                   // v8
import { DynamicFormContext } from "./DynamicFormContext";

/* ------------------------------------------------------------------ */
/* Types                                                              */
/* ------------------------------------------------------------------ */

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

  // match TagPicker: pass true to allow multiple selections
  multiselect?: boolean;
  disabled?: boolean;

  /* kept for NEW form compatibility – not used for Edit hydration */
  starterValue?:
    | { key: string | number; text: string }
    | { key: string | number; text: string }[];

  /* notify parent with SharePoint-style entities */
  onChange2?: (entities: PickerEntity[]) => void;

  /* optional knobs (defaults supplied) */
  principalType?: PrincipalType; // 1 = user only
  maxSuggestions?: number;       // default 5
  allowFreeText?: boolean;       // default false

  /* optional SPFx client + config for first-class POST */
  spHttpClient?: any;
  spHttpClientConfig?: any;
}

/* ------------------------------------------------------------------ */
/* Helpers / shared pieces                                            */
/* ------------------------------------------------------------------ */

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

/* ------------------------------------------------------------------ */
/* Component                                                           */
/* ------------------------------------------------------------------ */

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
    onChange2,
    principalType = 1,
    maxSuggestions = 5,
    allowFreeText = false,
    spHttpClient,
    spHttpClientConfig,
  } = props;

  const onChange = onChange2;

  const requiredEffective = (isRequired ?? isrequired2) ?? false;
  const isMulti = multiselect === true;

  // Explicit site URL (same pattern you used earlier)
  const webUrl =
    "https://amerihealthcaritas.sharepoint.com/sites/eokm";
  const apiUrl =
    `${webUrl}/_api/SP.UI.ApplicationPages.ClientPeoplePickerWebServiceInterface.clientPeoplePickerSearchUser`;

  const elemRef = React.useRef<HTMLDivElement | null>(null);

  // Normalize starterValue into ITag[] (only used for NEW form if provided)
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

  const [selectedTags, setSelectedTags] =
    React.useState<ITag[]>(normalizedStarter);
  const [lastResolved, setLastResolved] =
    React.useState<PickerEntity[]>([]);
  const [errorMsg, setErrorMsg] = React.useState<string>("");

  /* ---------- Global error handling (same style as TagPicker) ---------- */

  const reportError = React.useCallback(
    (msg: string) => {
      const targetId = `${id}Id`;
      setErrorMsg(msg || "");
      if (ctx && typeof ctx.GlobalErrorHandle === "function") {
        ctx.GlobalErrorHandle(targetId, msg || undefined);
      }
    },
    [ctx, id]
  );

  /* ---------- Register ref globally (GlobalRefs) ---------- */

  React.useEffect(() => {
    if (ctx && typeof ctx.GlobalRefs === "function") {
      ctx.GlobalRefs(
        elemRef.current !== null ? elemRef.current : undefined
      );
    }
  }, [ctx, elemRef]);

  /* ---------- Search (REST people API) ---------- */

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
          PrincipalSource: 1,       // All
          PrincipalType: principalType,
          QueryString: query,
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

  /* ---------- Change mapping back to entities + validation ---------- */

  const handleChange = React.useCallback(
    (items?: ITag[]) => {
      const rawNext = items ?? [];

      // Respect single vs multi
      const next = isMulti ? rawNext : rawNext.slice(-1);

      setSelectedTags(next);

      // Required-field validation
      if (requiredEffective && next.length === 0) {
        reportError("This field is required.");
      } else {
        reportError("");
      }

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

        // Free-text fallback (if allowed)
        if (allowFreeText) {
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
    [
      isMulti,
      requiredEffective,
      reportError,
      onChange,
      lastResolved,
      allowFreeText,
    ]
  );

  /* ---------- EDIT / VIEW FORM: hydrate PeoplePicker from ctx.FormData (SPUserID) ---------- */

  React.useEffect(() => {
    // Only run for EditForm(6) or ViewForm(4)
    if (!(ctx.FormMode === 4 || ctx.FormMode === 6)) {
      return;
    }

    const formData = ctx.FormData as any | undefined;
    if (!formData) return;

    const fieldInternalName = id;

    // Look at <InternalName>Id then <InternalName>StringId
    let rawValue: any = formData[`${fieldInternalName}Id`];
    if (rawValue === undefined) {
      rawValue = formData[`${fieldInternalName}StringId`];
    }
    if (rawValue === null || rawValue === undefined) return;

    // Normalize whatever SP stored into numeric SPUserID[]
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
              headers: {
                Accept: "application/json;odata=verbose",
              },
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
          console.error(
            "PeoplePicker getUserById error",
            err
          );
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

  /* ---------- NEW FORM: reset picker state so search works normally ---------- */

  React.useEffect(() => {
    if (ctx.FormMode === 8) {
      // New form – start clean (but honor any starterValue)
      setSelectedTags(normalizedStarter);
      setLastResolved([]);
      reportError("");
    }
  }, [ctx.FormMode, normalizedStarter, reportError]);

  /* ------------------------------------------------------------------ */
  /* Rendering                                                           */
  /* ------------------------------------------------------------------ */

  const requiredMsg =
    requiredEffective && selectedTags.length === 0
      ? "This field is required."
      : undefined;

  const finalValidationMsg = errorMsg || requiredMsg || undefined;
  const validationState = finalValidationMsg ? "error" : "none";

  const isDisabled = Boolean(submitting || disabled);

  return (
    <div
      ref={elemRef}
      className={className}
      // PeoplePicker itself doesn't track hidden state; if needed
      // the outer builder can wrap this component in a container.
    >
      <Field
        label={displayName}
        hint={description}
        validationMessage={finalValidationMsg}
        validationState={validationState}
      >
        <TagPicker
          className={className}
          disabled={isDisabled}
          // suggestions: respect single vs multi and current selection
          onResolveSuggestions={(filter, selected) => {
            const sel = selected ?? [];

            // Single-select: once one item is chosen, stop suggesting more
            if (!isMulti && sel.length >= 1) {
              return [];
            }

            if (
              !filter ||
              (sel.length &&
                sel.some(
                  (s) => String(s.key) === filter
                ))
            ) {
              return [];
            }

            return searchPeople(filter).then((tags) =>
              tags.filter(
                (t) =>
                  !(sel ?? []).some(
                    (s) =>
                      String(s.key) ===
                      String(t.key)
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
    </div>
  );
};

export default PeoplePicker;

