// PeoplePickerComponent.tsx
import * as React from "react";
import { Field } from "@fluentui/react-components"; // v9
import {
  TagPicker,
  ITag,
  IBasePickerSuggestionsProps,
} from "@fluentui/react"; // v8
import { DynamicFormContext } from "./DynamicFormContext";

/* ------------------------------------------------------------------ */
/* Types                                                               */
/* ------------------------------------------------------------------ */

export type PrincipalType = 0 | 1 | 2 | 4 | 8 | 15;

export interface PickerEntity {
  Key: string;
  DisplayText?: string;
  EntityType?: string;
  IsResolved?: boolean;
  EntityData?: any;
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
  multiselect?: boolean;
  disabled?: boolean;

  /** starterValue is only used for NEW forms – kept for builder compatibility */
  starterValue?:
    | { key: string | number; text: string }
    | { key: string | number; text: string }[];

  /** notify parent with SharePoint-style entities */
  onChange2?: (entities: PickerEntity[]) => void;

  /* optional knobs */
  principalType?: PrincipalType; // default 1 = User only
  maxSuggestions?: number; // default 5
  allowFreeText?: boolean; // default false

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

  const requiredEffective = (isRequired ?? isrequired2) ?? false;
  const isMulti = multiselect === true;

  // Explicit site URL + PeoplePicker service
  const webUrl = "https://amerihealthcaritas.sharepoint.com/sites/eokm";
  const apiUrl =
    `${webUrl}/_api/SP.UI.ApplicationPages.ClientPeoplePickerWebServiceInterface.ClientPeoplePickerSearchUser`;

  /* ------------------------------------------------------------------ */
  /* starterValue → initial tags (NEW form only)                        */
  /* ------------------------------------------------------------------ */

  const normalizedStarter: ITag[] = React.useMemo(() => {
    const baseArray: { key: string | number; text: string }[] =
      Array.isArray(starterValue)
        ? starterValue
        : starterValue
        ? [starterValue]
        : [];

    const arr = isMulti ? baseArray : baseArray.slice(-1);

    return arr.map((v) => ({
      key: String(v.key),
      name: v.text,
    }));
  }, [starterValue, isMulti]);

  /* ------------------------------------------------------------------ */
  /* Local state                                                        */
  /* ------------------------------------------------------------------ */

  const [selectedTags, setSelectedTags] =
    React.useState<ITag[]>(normalizedStarter);
  const [lastResolved, setLastResolved] =
    React.useState<PickerEntity[]>([]);
  const [displayOverride, setDisplayOverride] =
    React.useState<string>("");
  const [error, setError] = React.useState<string>("");

  // local flags for visibility / disabled styling
  const [isDisabledLocal, setIsDisabledLocal] =
    React.useState<boolean>(!!disabled);
  const [isHiddenLocal, setIsHiddenLocal] =
    React.useState<boolean>(false);

  const elemRef = React.useRef<HTMLInputElement | null>(null);

  /* ------------------------------------------------------------------ */
  /* Global error handler (TagPicker pattern)                           */
  /* ------------------------------------------------------------------ */

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
  }, [requiredEffective, selectedTags]);

  /* ------------------------------------------------------------------ */
  /* Commit numeric SPUserID(s) to GlobalFormData + GlobalRefs          */
  /* ------------------------------------------------------------------ */

  const commitValue = React.useCallback(() => {
    const err = validate();
    reportError(err);

    const targetId = `${id}Id`;

    // build lookup from lastResolved based on Key / AccountName / Email / DisplayText
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

    const ids: number[] = [];
    for (const t of selectedTags) {
      const lk = String(t.key).toLowerCase();
      const entity = resolvedByKey.get(lk);
      const spId =
        (entity as any)?.EntityData?.SPUserID ??
        (entity as any)?.EntityData2?.SPUserID ??
        (entity as any)?.Id;

      if (spId && !Number.isNaN(Number(spId))) {
        ids.push(Number(spId));
      }
    }

    if (isMulti) {
      ctx.GlobalFormData(targetId, ids.length === 0 ? [] : ids);
    } else {
      ctx.GlobalFormData(targetId, ids.length === 0 ? null : ids[0]);
    }

    const labels = selectedTags.map((t) => t.name);
    setDisplayOverride(labels.join("; "));

    ctx.GlobalRefs(
      elemRef.current !== null ? elemRef.current : undefined
    );
  }, [ctx, id, isMulti, lastResolved, selectedTags, reportError, validate]);

  const handleBlur = React.useCallback(() => {
    commitValue();
  }, [commitValue]);

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
          PrincipalSource: 1,
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
          setLastResolved(entities);
          return entities.map(toTag);
        }

        // Fallback classic fetch with request digest
        const digest = (
          document.getElementById(
            "__REQUESTDIGEST"
          ) as HTMLInputElement | null
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
            "PeoplePicker fetch error",
            resp.status,
            resp.statusText,
            txt
          );
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
  /* TagPicker onChange → back to entities                               */
  /* ------------------------------------------------------------------ */

  const handleChange = React.useCallback(
    (items?: ITag[] | null) => {
      let next = items ?? [];

      // enforce single-select if not multiselect
      if (!isMulti && next.length > 1) {
        next = next.slice(-1);
      }

      setSelectedTags(next);

      if (!onChange2) return;

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

        if (allowFreeText) {
          result.push({
            Key: String(t.key),
            DisplayText: t.name,
            IsResolved: false,
            EntityType: "User",
            EntityData2: /@/.test(String(t.key))
              ? { Email: String(t.key) }
              : {},
          });
        }
      }

      onChange2(result);
    },
    [allowFreeText, isMulti, lastResolved, onChange2]
  );

  /* ------------------------------------------------------------------ */
  /* EDIT / VIEW FORM: hydrate from ctx.FormData (SPUserID)             */
  /* ------------------------------------------------------------------ */

  React.useEffect(() => {
    if (!(ctx.FormMode === 4 || ctx.FormMode === 6)) return; // 4=view, 6=edit

    const fieldInternalName = id;
    const formData = ctx.FormData as any | undefined;
    if (!formData) return;

    let rawValue: any = formData[fieldInternalName];

    if (rawValue === undefined) {
      const idProp = `${fieldInternalName}Id`;
      const stringIdProp = `${fieldInternalName}StringId`;
      rawValue = formData[idProp] ?? formData[stringIdProp];
    }

    if (rawValue === null || rawValue === undefined) return;

    const collectIds = (value: any): number[] => {
      if (value === null || value === undefined) return [];

      if (Array.isArray(value)) {
        const ids: number[] = [];
        for (const v of value) {
          if (typeof v === "object" && "Id" in v) {
            ids.push(Number((v as any).Id));
          } else {
            ids.push(Number(v));
          }
        }
        return ids.filter((n) => !Number.isNaN(n));
      }

      const str = String(value);
      const parts = str.split(/[;,#]/);
      return parts
        .map((p) => Number(p.trim()))
        .filter((n) => !Number.isNaN(n));
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
          console.error("PeoplePicker getUserById error", err);
        }
      }

      if (!hydrated.length) return;

      setLastResolved(hydrated);
      const tags = hydrated.map(toTag);
      setSelectedTags(tags);

      if (onChange2) {
        onChange2(hydrated);
      }
    })();

    return () => abort.abort();
    // run once for this field; don't depend on ctx to avoid loops
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [id, webUrl, onChange2]);

  /* ------------------------------------------------------------------ */
  /* NEW FORM: reset state so search behaves normally                   */
  /* ------------------------------------------------------------------ */

  React.useEffect(() => {
    if (ctx.FormMode === 8) {
      // New form – initialise with starter tags, clear errors
      setSelectedTags(normalizedStarter);
      setLastResolved([]);
      setDisplayOverride("");
      setError("");
      setIsDisabledLocal(!!disabled);
      setIsHiddenLocal(false);

      ctx.GlobalRefs(
        elemRef.current !== null ? elemRef.current : undefined
      );
    }
    // IMPORTANT: no `ctx` object in deps → avoids loop when context updates
  }, [ctx.FormMode, normalizedStarter, disabled, ctx.GlobalRefs]);

  /* ------------------------------------------------------------------ */
  /* Rendering                                                          */
  /* ------------------------------------------------------------------ */

  const requiredMsg =
    requiredEffective && selectedTags.length === 0
      ? "This field is required."
      : undefined;

  const hasError = error || requiredMsg;
  const disabledFinal = Boolean(
    submitting || disabled || isDisabledLocal
  );

  return (
    <Field
      label={displayName}
      hint={description}
      validationMessage={hasError}
      validationState={hasError ? "error" : "none"}
      style={{ display: isHiddenLocal ? "none" : "block" }}
    >
      <TagPicker
        className={className}
        disabled={disabledFinal}
        onResolveSuggestions={(filter, selected) => {
          if (!filter || (selected ?? []).length >= maxSuggestions) {
            return [];
          }
          return searchPeople(filter).then((tags) =>
            tags.filter(
              (t) =>
                (selected ?? []).every(
                  (s) => String(s.key) !== String(t.key)
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
          onBlur: handleBlur,
          // TagPicker's input is not strongly typed, so cast
          ref: elemRef as any,
        }}
      />

      {displayOverride && (
        <div style={{ marginTop: 4, fontSize: 12, opacity: 0.7 }}>
          {displayOverride}
        </div>
      )}
    </Field>
  );
};

export default PeoplePicker;




