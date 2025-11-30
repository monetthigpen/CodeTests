// PeoplePickerComponent.tsx

import * as React from "react";
import { Field } from "@fluentui/react-components";            // v9
import {
  TagPicker,
  ITag,
  IBasePickerSuggestionsProps,
} from "@fluentui/react";                                     // v8
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
  displayName2?: string;
  className2?: string;
  description2?: string;
  placeholder2?: string;

  isRequired2?: boolean;
  isrequired2?: boolean;
  submitting2?: boolean;

  multiselect2?: boolean;
  disabled2?: boolean;

  /** optional default(s); kept for compatibility but NOT used for Edit hydrations */
  starterValue2?:
    | { key: string | number; text: string }
    | { key: string | number; text: string }[];

  /** notify parent with SharePoint-style entities */
  onChange2?: (entities: PickerEntity[]) => void;

  /* optional knobs (defaults supplied) */
  principalType2?: PrincipalType; // 1 = user only
  maxSuggestions2?: number;       // default 5
  allowFreeText2?: boolean;       // default false

  /* optional SPFx client + config for first-class POST */
  spHttpClient2?: any;
  spHttpClientConfig2?: any;

  /* tolerate unknown props from builder */
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
/* Component                                                           */
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
    submitting2,
    multiselect2,
    disabled2,
    starterValue2,
    onChange2,
    principalType2 = 1,
    maxSuggestions2 = 5,
    allowFreeText2 = false,
    spHttpClient2,
    spHttpClientConfig2,
  } = props;

  // Normalize builder props
  const displayName = displayName2 ?? id;
  const className = className2 ?? "elementsWidth";
  const description = description2;
  const placeholder = placeholder2;
  const requiredEffective = (isRequired2 ?? isrequired2) ?? false;
  const submitting = !!submitting2;
  const isMulti = multiselect2 === true;
  const initiallyDisabled = !!disabled2;
  const starterValue = starterValue2;
  const principalType = principalType2;
  const maxSuggestions = maxSuggestions2 ?? 5;
  const allowFreeText = allowFreeText2 === true;
  const spHttpClient = spHttpClient2;
  const spHttpClientConfig = spHttpClientConfig2;
  const onChange = onChange2;

  // Explicit site URL (same pattern you used earlier)
  const webUrl = "https://amerihealthcaritas.sharepoint.com/sites/eokm";
  const apiUrl = `${webUrl}/_api/SP.UI.ApplicationPages.ClientPeoplePickerWebServiceInterface.clientPeoplePickerSearchUser`;

  // ----- Normalize starterValue into ITag[] (used only for NEW forms if provided)
  const starterArray: { key: string | number; text: string }[] =
    starterValue == null
      ? []
      : Array.isArray(starterValue)
      ? starterValue
      : [starterValue];

  const normalizedStarter: ITag[] = (isMulti ? starterArray : starterArray.slice(-1)).map(
    (v) => ({
      key: String(v.key),
      name: v.text,
    })
  );

  // ----- Local state -------------------------------------------------

  const [selectedTags, setSelectedTags] = React.useState<ITag[]>(normalizedStarter);
  const [lastResolved, setLastResolved] = React.useState<PickerEntity[]>([]);
  const [isDisabled, setIsDisabled] = React.useState<boolean>(initiallyDisabled);
  const [defaultDisable, setDefaultDisable] = React.useState<boolean>(false);
  const [isHidden] = React.useState<boolean>(false); // currently no dynamic hide rules
  const [error, setError] = React.useState<string>("");

  const elemRef = React.useRef<HTMLDivElement | null>(null);

  /* ------------------------ Global error helper --------------------- */

  const reportError = React.useCallback(
    (msg: string) => {
      const targetId = `${id}Id`;
      setError(msg || "");
      ctx.GlobalErrorHandle(targetId, msg || undefined);
    },
    [ctx, id]
  );

  const validate = React.useCallback((): string => {
    if (requiredEffective && selectedTags.length === 0) {
      return "This field is required.";
    }
    return "";
  }, [requiredEffective, selectedTags.length]);

  /* ------------------------ Search (REST people API) ---------------- */

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
            console.error("PeoplePicker spHttpClient error:", resp.status, resp.statusText, txt);
            return [];
          }

          const data: any = await resp.json();
          const raw = data.d?.ClientPeoplePickerSearchUser ?? "[]";
          const entities: PickerEntity[] = JSON.parse(raw).map((e: any) => ({
            Key: String(e.EntityData?.SPUserID ?? e.Key),
            DisplayText: e.DisplayText,
            EntityType: e.EntityType,
            IsResolved: e.IsResolved,
            EntityData2: {
              Email: e.EntityData?.Email,
              AccountName: e.EntityData?.AccountName,
              Title: e.EntityData?.Title,
              Department2: e.EntityData?.Department ?? "",
            },
          }));

          setLastResolved(entities);
          return entities.map(toTag);
        }

        // Fallback: classic fetch with request digest
        const digest = (document.getElementById("__REQUESTDIGEST") as HTMLInputElement | null)
          ?.value;

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
        const entities: PickerEntity[] = JSON.parse(raw).map((e: any) => ({
          Key: String(e.EntityData?.SPUserID ?? e.Key),
          DisplayText: e.DisplayText,
          EntityType: e.EntityType,
          IsResolved: e.IsResolved,
          EntityData2: {
            Email: e.EntityData?.Email,
            AccountName: e.EntityData?.AccountName,
            Title: e.EntityData?.Title,
            Department2: e.EntityData?.Department ?? "",
          },
        }));

        setLastResolved(entities);
        return entities.map(toTag);
      } catch (e) {
        console.error("PeoplePicker exception:", e);
        return [];
      }
    },
    [apiUrl, isMulti, maxSuggestions, principalType, spHttpClient, spHttpClientConfig]
  );

  /* ------------------------ Change mapping back to entities --------- */

  const handleChange = React.useCallback(
    (items?: ITag[]) => {
      const next = items ?? [];
      setSelectedTags(next);

      // Build lookup from last resolved entities by Key / AccountName / Email
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
        const match = resolvedByKey.get(lk);
        if (match) {
          result.push(match);
          continue;
        }

        // synthesize minimal entity from free text (if allowed)
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

      // send to parent
      if (onChange) {
        onChange(result);
      }

      // store numeric SPUserID(s) in GlobalFormData
      const targetId = `${id}Id`;
      const ids = result
        .map((e) => Number(e.Key))
        .filter((n) => !Number.isNaN(n));

      const storedValue = isMulti ? ids : ids[0] ?? null;
      ctx.GlobalFormData(targetId, storedValue);

      // validation + global error
      const msg = validate();
      reportError(msg);

      // make sure this element is registered as a GlobalRef
      ctx.GlobalRefs(elemRef.current !== null ? elemRef.current : undefined);
    },
    [allowFreeText, ctx, id, isMulti, lastResolved, onChange, reportError, validate]
  );

  /* ------------------------ EDIT / VIEW FORM: hydrate from ctx.FormData ---- */

  React.useEffect(() => {
    // Only run for EditForm(6) or ViewForm(4)
    if (!(ctx.FormMode === 4 || ctx.FormMode === 6)) {
      return;
    }

    const formData = ctx.FormData as any | undefined;
    if (!formData) return;

    // Form data can store variants:
    //   <InternalName>
    //   <InternalName>Id
    //   <InternalName>StringId
    const fieldInternalName = id;
    let rawValue: any = formData[fieldInternalName];

    if (rawValue === undefined) {
      const idProp = `${fieldInternalName}Id`;
      const stringIdProp = `${fieldInternalName}StringId`;
      rawValue = formData[idProp] ?? formData[stringIdProp];
    }

    if (rawValue === null || rawValue === undefined) return;

    // normalize whatever SP stored into numeric SPUserID[]
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
          const resp = await fetch(`${webUrl}/_api/web/getUserById(${userId})`, {
            method: "GET",
            headers: { Accept: "application/json;odata=verbose" },
            signal: abort.signal,
          });

          if (!resp.ok) {
            console.warn("PeoplePicker getUserById failed", userId, resp.status, resp.statusText);
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

      const msg = validate();
      reportError(msg);

      const targetId = `${id}Id`;
      const ids = hydrated
        .map((e) => Number(e.Key))
        .filter((n) => !Number.isNaN(n));
      const storedValue = isMulti ? ids : ids[0] ?? null;
      ctx.GlobalFormData(targetId, storedValue);

      ctx.GlobalRefs(elemRef.current !== null ? elemRef.current : undefined);
    })();

    return () => abort.abort();
  }, [ctx.FormMode, ctx.FormData, id, isMulti, reportError, validate, webUrl, ctx]);

  /* ------------------------ NEW FORM: make sure we start clean so search works --- */

  React.useEffect(() => {
    if (ctx.FormMode === 8) {
      setSelectedTags(normalizedStarter);
      setLastResolved([]);
      setError("");
      ctx.GlobalRefs(elemRef.current !== null ? elemRef.current : undefined);
    }
  }, [ctx.FormMode, normalizedStarter, ctx]);

  /* ------------------------ Always register this field as a GlobalRef ----------- */

  React.useEffect(() => {
    ctx.GlobalRefs(elemRef.current !== null ? elemRef.current : undefined);
  }, [ctx]);

  /* ------------------------ Rendering ------------------------------------------- */

  const hasError = !!error;
  const requiredMsg = hasError ? error : undefined;
  const effectiveDisabled = submitting || defaultDisable || isDisabled;

  return (
    <div
      ref={elemRef}
      style={{ display: isHidden ? "none" : "block" }}
      className={className}
      data-disabled={effectiveDisabled ? "true" : undefined}
    >
      <Field
        label={displayName}
        hint={description}
        validationMessage={requiredMsg}
        validationState={hasError ? "error" : "none"}
      >
        <TagPicker
          onResolveSuggestions={(filter, selected) => {
            const already = (selected ?? []).map((t) => String(t.key));
            return searchPeople(filter || "").then((tags) =>
              tags.filter((t) => !already.some((k) => k === String(t.key)))
            );
          }}
          getTextFromItem={(t) => t.name}
          selectedItems={selectedTags}
          onChange={handleChange}
          pickerSuggestionsProps={suggestionsProps}
          inputProps={{ placeholder: placeholder ?? "Search people…" }}
          disabled={effectiveDisabled}
        />
      </Field>
    </div>
  );
};

export default PeoplePicker;

