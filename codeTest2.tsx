// PeoplePickerComponent.tsx
import * as React from "react";
import { Field } from "@fluentui/react-components";             // v9
import {
  TagPicker,
  ITag,
  IBasePickerSuggestionsProps,
} from "@fluentui/react";                                      // v8
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
    SPUserID?: number;          // for edit hydrations & GlobalFormData
  };
}

export interface PeoplePickerProps {
  id: string;
  displayName2?: string;
  className2?: string;
  description2?: string;
  placeholder2?: string;

  isRequired2?: boolean;
  isrequired2?: boolean;
  submitting?: boolean;

  /** match TagPicker: pass true to allow multiple selections */
  multiselect?: boolean;
  disabled?: boolean;

  /** optional starter value: NOT used for Edit hydrations, only NEW */
  starterValue?:
    | { key: string | number; text: string }
    | { key: string | number; text: string }[];

  /** notify parent with SharePoint-style entities */
  onChange2?: (entities: PickerEntity[]) => void;

  /* optional knobs (defaults supplied) */
  principalType?: PrincipalType;   // 1 = user only
  maxSuggestions?: number;         // default 5
  allowFreeText?: boolean;         // default false

  /* optional SPFx client + config for first-class POST */
  spHttpClient2?: any;
  spHttpClientConfig2?: any;

  /* tolerate unknown props from builder */
  [key: string]: any;
}

/* ------------------------------------------------------------------ */
/* Helpers / shared pieces                                            */
/* ------------------------------------------------------------------ */

const REQUIRED_MSG = "This field is required.";

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
/* Component                                                          */
/* ------------------------------------------------------------------ */

const PeoplePicker: React.FC<PeoplePickerProps> = (props) => {
  const ctx = React.useContext(DynamicFormContext);

  const {
    id,
    displayName2,
    className2,
    description2,
    placeholder2,
    isRequired2,
    isrequired2,
    submitting,
    multiselect,
    disabled,
    starterValue,
    onChange2,
    principalType = 1,
    maxSuggestions = 5,
    allowFreeText = false,
    spHttpClient2,
    spHttpClientConfig2,
  } = props;

  /* effective, TagPicker-style props */
  const displayName = displayName2 ?? (props as any).displayName ?? id;
  const className = className2 ?? (props as any).className;
  const description = description2 ?? (props as any).description;
  const placeholder = placeholder2 ?? (props as any).placeholder;
  const isRequired = isRequired2 ?? isrequired2 ?? false;

  const requiredEffective = !!isRequired;
  const isMulti = multiselect === true;
  const isSubmitting = !!submitting;

  // Explicit site URL (same pattern as in your screenshots)
  const webUrl =
    "https://amerihealthcaritas.sharepoint.com/sites/eokm";
  const apiUrl =
    `${webUrl}/_api/SP.UI.ApplicationPages.ClientPeoplePickerWebServiceInterface.clientPeoplePickerSearchUser`;

  /* -------------------------- React state -------------------------- */

  const [selectedTags, setSelectedTags] = React.useState<ITag[]>([]);
  const [lastResolved, setLastResolved] = React.useState<PickerEntity[]>([]);
  const [isDisabled, setIsDisabled] = React.useState<boolean>(!!disabled);
  const [error, setError] = React.useState<string>("");
  const [displayOverride, setDisplayOverride] = React.useState<string>("");

  const elemRef = React.useRef<HTMLDivElement | null>(null);

  /* ------------------------- Global helpers ------------------------ */

  // Register DOM element with global refs (for focus/scroll, etc.)
  React.useEffect(() => {
    if (ctx && ctx.GlobalRefs) {
      ctx.GlobalRefs(
        elemRef.current !== null ? elemRef.current : undefined,
        id
      );
    }
  }, [ctx, id]);

  // Report validation error through global handler
  const reportError = React.useCallback(
    (msg: string) => {
      const targetId = `${id}Id`;
      const next = msg || "";
      setError(next);
      if (ctx && ctx.GlobalErrorHandle) {
        ctx.GlobalErrorHandle(targetId, next || undefined);
      }
    },
    [ctx, id]
  );

  // Validate required-ness like TagPicker
  const validate = React.useCallback((): string => {
    if (requiredEffective && selectedTags.length === 0) {
      return REQUIRED_MSG;
    }
    return "";
  }, [requiredEffective, selectedTags.length]);

  // Commit a full set of resolved entities into GlobalFormData + fire onChange2
  const commitValue = React.useCallback(
    (entities: PickerEntity[]) => {
      // Required validation
      const err = validate();
      reportError(err);

      // Build numeric SPUserID list (fall back to numeric Key)
      const ids: number[] = [];
      for (const e of entities) {
        let idVal: number | undefined;
        if (typeof e.EntityData2?.SPUserID === "number") {
          idVal = e.EntityData2.SPUserID;
        } else {
          const numKey = Number(e.Key);
          if (!Number.isNaN(numKey)) {
            idVal = numKey;
          }
        }
        if (typeof idVal === "number" && !Number.isNaN(idVal)) {
          ids.push(idVal);
        }
      }

      const targetId = `${id}Id`;
      if (ctx && ctx.GlobalFormData) {
        if (isMulti) {
          ctx.GlobalFormData(targetId, ids.length === 0 ? [] : ids);
        } else {
          ctx.GlobalFormData(
            targetId,
            ids.length === 0 ? null : ids[0]
          );
        }
      }

      if (onChange2) {
        onChange2(entities);
      }

      // Display text (similar to TagPicker's displayOverride)
      const labels = entities.map((e) => e.DisplayText ?? e.Key);
      setDisplayOverride(labels.join("; "));
    },
    [ctx, id, isMulti, onChange2, validate, reportError]
  );

  /* ---------------------- Normalize starterValue ------------------- */
  // Only used for NEW forms if builder provides; edit hydration uses ctx.FormData.

  const starterArray: { key: string | number; text: string }[] =
    starterValue == null
      ? []
      : Array.isArray(starterValue)
      ? starterValue
      : [starterValue];

  const normalizedStarter: ITag[] = (isMulti
    ? starterArray
    : starterArray.slice(-1)
  ).map((v) => ({
    key: String(v.key),
    name: v.text,
  }));

  /* -------------------------- Search (REST) ------------------------ */

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
        // Prefer SPFx HttpClient if provided (like your TagPicker)
        if (spHttpClient2 && spHttpClientConfig2) {
          const resp = await spHttpClient2.post(
            apiUrl,
            spHttpClientConfig2,
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
              "PeoplePicker spHttpClient error",
              resp.status,
              resp.statusText,
              txt
            );
            return [];
          }

          const data: any = await resp.json();
          const raw = data.d?.ClientPeoplePickerSearchUser ?? "[]";
          const entities: PickerEntity[] = JSON.parse(raw);
          // Capture SPUserID if present
          entities.forEach((e: any) => {
            const ed = (e.EntityData || e.EntityData2) as any;
            if (ed && typeof ed.SPUserID === "number") {
              e.EntityData2 = {
                ...(e.EntityData2 || {}),
                SPUserID: ed.SPUserID,
              };
            }
          });
          setLastResolved(entities);
          return entities.map(toTag);
        }

        // Fallback: classic fetch with request digest
        const digest = (
          document.getElementById("__REQUESTDIGEST") as
            | HTMLInputElement
            | null
        )?.value;

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
        entities.forEach((e: any) => {
          const ed = (e.EntityData || e.EntityData2) as any;
          if (ed && typeof ed.SPUserID === "number") {
            e.EntityData2 = {
              ...(e.EntityData2 || {}),
              SPUserID: ed.SPUserID,
            };
          }
        });
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
      spHttpClient2,
      spHttpClientConfig2,
    ]
  );

  /* ------------------ Tag -> PickerEntity mapping ------------------ */

  const handleChange = React.useCallback(
    (items?: ITag[]) => {
      const next = items ?? [];
      setSelectedTags(next);

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
              Title: t.name,
              Department2: "",
            },
          });
        }
      }

      commitValue(result);
    },
    [allowFreeText, commitValue, lastResolved]
  );

  /* ---------------- EDIT / VIEW: hydrate from ctx.FormData --------- */
  // SPUserID values pulled from FormData then resolved via REST

  React.useEffect(() => {
    // Only run for Edit(6) or View(4)
    if (!(ctx.FormMode === 4 || ctx.FormMode === 6)) {
      return;
    }

    const formData = ctx.FormData as any | undefined;
    if (!formData) return;

    const fieldInternalName = id;

    // Try <InternalName>Id, then <InternalName>StringId, then <InternalName>
    let rawValue: any = formData[fieldInternalName];
    if (rawValue === undefined) {
      const idProp = `${fieldInternalName}Id`;
      const stringIdProp = `${fieldInternalName}StringId`;
      rawValue = formData[idProp] ?? formData[stringIdProp];
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
              SPUserID: u.Id,
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

      // push into GlobalFormData + onChange2
      commitValue(hydrated);
    })();

    return () => abort.abort();
  }, [ctx.FormMode, ctx.FormData, id, commitValue, webUrl]);

  /* ---------------- NEW FORM: reset state clean -------------------- */

  React.useEffect(() => {
    if (ctx.FormMode === 8) {
      // New form: start from starterValue only & clear resolved cache
      setSelectedTags(normalizedStarter);
      setLastResolved([]);
      setError("");
      if (ctx && ctx.GlobalFormData) {
        const targetId = `${id}Id`;
        const starterIds: number[] = [];
        for (const t of normalizedStarter) {
          const n = Number(t.key);
          if (!Number.isNaN(n)) starterIds.push(n);
        }
        if (isMulti) {
          ctx.GlobalFormData(
            targetId,
            starterIds.length === 0 ? [] : starterIds
          );
        } else {
          ctx.GlobalFormData(
            targetId,
            starterIds.length === 0 ? null : starterIds[0]
          );
        }
      }
    }
  }, [ctx, ctx.FormMode, id, isMulti, normalizedStarter]);

  /* ---------------------- Submitting / disable --------------------- */

  React.useEffect(() => {
    if (isSubmitting === false) {
      setIsDisabled(!!disabled);
    } else {
      setIsDisabled(true);
    }
  }, [isSubmitting, disabled]);

  /* ----------------------------- Render ---------------------------- */

  const hasError = !!error || (requiredEffective && selectedTags.length === 0);
  const validationMsg =
    hasError ? error || REQUIRED_MSG : undefined;

  return (
    <div
      ref={elemRef}
      style={{ display: ctx?.isHidden ? "none" : "block" }}
      className={className}
    >
      <Field
        label={displayName}
        hint={description}
        validationMessage={validationMsg}
        validationState={hasError ? "error" : "none"}
      >
        <TagPicker
          className={className}
          disabled={isDisabled}
          onResolveSuggestions={(filter, selected) => {
            const sel = selected ?? [];
            if (!filter || !filter.trim()) return [];
            return searchPeople(filter).then((tags) =>
              tags.filter(
                (t) =>
                  !sel.some(
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
    </div>
  );
};

export default PeoplePicker;
