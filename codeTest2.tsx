// PeoplePickerComponent.tsx
import * as React from "react";
import { Field } from "@fluentui/react-components";            // v9
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
  Key: string;             // SPUserID as string
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

  /** match TagPicker: pass true to allow multiple selections */
  multiselect?: boolean;
  disabled?: boolean;

  /** optional default(s); kept for compatibility but NOT used for Edit hydrations */
  starterValue?:
    | { key: string | number; text: string }
    | { key: string | number; text: string }[];

  /** notify parent with SharePoint-style entities */
  onChange2?: (entities: PickerEntity[]) => void;

  /** optional knobs (defaults supplied) */
  principalType?: PrincipalType;   // 1 = user only
  maxSuggestions?: number;         // default 5
  allowFreeText?: boolean;         // default false

  /** optional SPFx client + config for first-class POST */
  spHttpClient?: any;
  spHttpClientConfig?: any;

  /** tolerate unknown props from builder */
  [key: string]: any;
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
/* Component                                                          */
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

  const elemRef = React.useRef<HTMLDivElement | null>(null);

  // local error message (mirrors GlobalErrorHandle like TagPicker)
  const [error, setError] = React.useState<string>("");

  const requiredEffective = (isRequired ?? isrequired2) ?? false;
  const isMulti = multiselect === true;

  // Explicit site URL – same pattern you used earlier
  const webUrl = "https://amerihealthcaritas.sharepoint.com/sites/eokm";
  const apiUrl = `${webUrl}/_api/SP.UI.ApplicationPages.ClientPeoplePickerWebServiceInterface.clientPeoplePickerSearchUser`;

  // ---- Normalize starterValue into ITag[] (used only for NEW forms if provided) ----
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

  /* ------------------------------------------------------------------ */
  /* Global error helper + refs (same pattern as TagPickerComponent)    */
  /* ------------------------------------------------------------------ */

  const reportError = React.useCallback(
    (msg: string) => {
      const targetId = `${id}Id`;
      setError(msg || "");
      if (ctx.GlobalErrorHandle) {
        ctx.GlobalErrorHandle(targetId, msg || undefined);
      }
    },
    [ctx, id]
  );

  // Register DOM ref globally (focus / validation)
  React.useEffect(() => {
    if (ctx.GlobalRefs) {
      ctx.GlobalRefs(elemRef.current ?? undefined);
    }
  }, [ctx]);

  /* ------------------------------------------------------------------ */
  /* Search (REST people API)                                           */
  /* ------------------------------------------------------------------ */

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
            console.error("PeoplePicker spHttpClient error", resp.status, resp.statusText, txt);
            return [];
          }

          const data: any = await resp.json();
          const raw = data.d?.ClientPeoplePickerSearchUser ?? "[]";
          const entities: PickerEntity[] = JSON.parse(raw);
          setLastResolved(entities);
          return entities.map(toTag);
        }

        // Fallback: classic fetch with request digest
        const digest = (document.getElementById("__REQUESTDIGEST") as HTMLInputElement | null)?.value;

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
          console.error("PeoplePicker fetch error:", resp.status, resp.statusText, txt);
          return [];
        }

        const json: any = await resp.json();
        const raw = json.d?.ClientPeoplePickerSearchUser ?? "[]";
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

  /* ------------------------------------------------------------------ */
  /* Change mapping back to entities + GlobalFormData                   */
  /* ------------------------------------------------------------------ */

  const handleChange = React.useCallback(
    (items?: ITag[]) => {
      const next = items ?? [];
      setSelectedTags(next);

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

      const result: PickerEntity[] = [];

      for (const t of next) {
        const lk = String(t.key).toLowerCase();
        const hit = resolvedByKey.get(lk);
        if (hit) {
          result.push(hit);
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

      // Push up to parent (SharePoint-style entities)
      if (onChange2) {
        onChange2(result);
      }

      // ----- GlobalFormData: store numeric SPUserID(s), like TagPicker does -----
      if (ctx.GlobalFormData) {
        const targetId = `${id}Id`;
        const numericIds = result
          .map((e) => Number(e.Key))
          .filter((n) => !Number.isNaN(n));

        if (isMulti) {
          ctx.GlobalFormData(targetId, numericIds.length === 0 ? [] : numericIds);
        } else {
          ctx.GlobalFormData(
            targetId,
            numericIds.length === 0 ? null : numericIds[0]
          );
        }
      }

      // Validation + GlobalError
      if (requiredEffective) {
        const msg =
          result.length === 0 ? "This field is required." : "";
        reportError(msg);
      }
    },
    [lastResolved, allowFreeText, onChange2, ctx, id, isMulti, requiredEffective, reportError]
  );

  /* ------------------------------------------------------------------ */
  /* EDIT / VIEW: hydrate PeoplePicker from ctx.FormData (SPUserID)     */
  /* ------------------------------------------------------------------ */

  React.useEffect(() => {
    // Only run for EditForm(6) or ViewForm(4)
    if (!(ctx.FormMode === 4 || ctx.FormMode === 6)) {
      return;
    }

    // If we already have something, don't re-hydrate
    if (lastResolved.length > 0 || selectedTags.length > 0) {
      return;
    }

    const formData = ctx.FormData as any | undefined;
    if (!formData) return;

    const fieldInternalName = id;

    // Look at <InternalName>, then <InternalName>Id, then <InternalName>StringId>
    let rawValue = formData[fieldInternalName];
    if (rawValue === undefined) {
      const idProp = `${fieldInternalName}Id`;
      const stringIdProp = `${fieldInternalName}StringId`;
      rawValue = formData[idProp] ?? formData[stringIdProp];
    }

    if (rawValue == null) return;

    const collectIds = (value: any): number[] => {
      if (value == null) return [];

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

      // NOTE: no onChange2() or GlobalFormData here – ctx.FormData already
      // contains the values and we don't want a feedback loop.
    })();

    return () => abort.abort();
  }, [ctx.FormMode, ctx.FormData, id, webUrl, lastResolved.length, selectedTags.length]);

  /* ------------------------------------------------------------------ */
  /* NEW FORM: reset picker so search behaves normally                  */
  /* ------------------------------------------------------------------ */

  React.useEffect(() => {
    if (ctx.FormMode === 8) {
      // New form – start clean
      setSelectedTags([]);
      setLastResolved([]);
      setError("");
    }
  }, [ctx.FormMode]);

  /* ------------------------------------------------------------------ */
  /* Rendering                                                          */
  /* ------------------------------------------------------------------ */

  const requiredMsg =
    requiredEffective && selectedTags.length === 0
      ? "This field is required."
      : undefined;

  const isDisabled = Boolean(submitting || disabled);

  const validationMsg = error || requiredMsg;
  const validationState = validationMsg ? "error" : "none";

  const picker = (
    <TagPicker
      onResolveSuggestions={(filter, selected) => {
        const sel = selected ?? [];
        if (!filter || (sel.length && sel.some((s) => String(s.key) === filter))) {
          return [];
        }
        return searchPeople(filter).then((tags) =>
          tags.filter(
            (t) =>
              !(sel ?? []).some((s) => String(s.key) === String(t.key))
          )
        );
      }}
      selectedItems={selectedTags}
      onChange={handleChange}
      pickerSuggestionsProps={suggestionsProps}
      inputProps={{ placeholder: placeholder ?? "Search people…" }}
      className={className}
      disabled={isDisabled}
    />
  );

  return displayName ? (
    <Field
      ref={elemRef as any}
      label={displayName}
      hint={description}
      validationMessage={validationMsg}
      validationState={validationState as any}
    >
      {picker}
    </Field>
  ) : (
    <div ref={elemRef}>{picker}</div>
  );
};

export default PeoplePicker;


