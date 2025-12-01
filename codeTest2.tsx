// PeoplePickerComponent.tsx
import * as React from "react";
import { Field } from "@fluentui/react-components";          // v9
import {
  TagPicker,
  ITag,
  IBasePickerSuggestionsProps
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
  isrequired2?: boolean;             // builder sometimes passes this – we normalize
  submitting?: boolean;

  multiselect?: boolean;
  disabled?: boolean;

  /** for NEW form only; Edit/View are hydrated from ctx.FormData */
  starterValue?:
    | { key: string | number; text: string }
    | { key: string | number; text: string }[];

  /** Notify parent with resolved SharePoint-style entities */
  onChange2?: (entities: PickerEntity[]) => void;

  /* optional knobs (defaults supplied below) */
  principalType?: PrincipalType;     // 1 = Users only, etc.
  maxSuggestions?: number;
  allowFreeText?: boolean;

  /* optional SPFx client, for first-class POST */
  spHttpClient?: any;
  spHttpClientConfig?: any;

  /* tolerant: extra values from builder */
  [key: string]: any;
}

/* ------------------------------------------------------------------ */
/* Shared helpers                                                     */
/* ------------------------------------------------------------------ */

const suggestionsProps: IBasePickerSuggestionsProps = {
  suggestionsHeaderText: "People",
  noResultsFoundText: "No results",
  resultsMaximumNumber: 5
};

/** Build an ITag from a people entity – never return undefined keys. */
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
    name: String(rawName)
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
    spHttpClientConfig
  } = props;

  const requiredEffective = (isRequired ?? isrequired2) ?? false;
  const isMulti = multiselect === true;

  // ----- site + API URLs -----
  const webUrl = "https://amerihealthcaritas.sharepoint.com/sites/eokm";
  const apiUrl =
    `${webUrl}/_api/SP.UI.ApplicationPages.ClientPeoplePickerWebServiceInterface.clientPeoplePickerSearchUser`;

  // -------------------------------------------------------------------
  // Global UI state (mirrors other controls – disabled, hidden, etc.)
  // -------------------------------------------------------------------
  const [isDisabled, setIsDisabled] = React.useState<boolean>(!!disabled);
  const [defaultDisable, setDefaultDisable] = React.useState<boolean>(false);
  const [isHidden, setIsHidden] = React.useState<boolean>(false);
  const [touched, setTouched] = React.useState<boolean>(false);
  const [displayOverride, setDisplayOverride] = React.useState<string>("");
  const [errorMsg, setErrorMsg] = React.useState<string>("");

  // Ref used for GlobalRefs – put it on an outer <div>, not on TagPicker
  const elemRef = React.useRef<HTMLDivElement | null>(null);

  // -------------------------------------------------------------------
  // People-picker specific state
  // -------------------------------------------------------------------
  const [selectedTags, setSelectedTags] = React.useState<ITag[]>([]);
  const [lastResolved, setLastResolved] = React.useState<PickerEntity[]>([]);

  // Normalized starter value – used only for NEW forms
  const starterArray: { key: string | number; text: string }[] = React.useMemo(
    () =>
      starterValue == null
        ? []
        : Array.isArray(starterValue)
        ? starterValue
        : [starterValue],
    [starterValue]
  );

  const normalizedStarter: ITag[] = React.useMemo(
    () =>
      starterArray.map((v) => ({
        key: String(v.key),
        name: v.text
      })),
    [starterArray]
  );

  // -------------------------------------------------------------------
  // Error reporting (GlobalErrorHandle) – follows TagPicker pattern
  // -------------------------------------------------------------------
  const reportError = React.useCallback(
    (msg: string) => {
      const targetId = `${id}Id`;
      setErrorMsg(msg || "");
      if (ctx.GlobalErrorHandle) {
        ctx.GlobalErrorHandle(targetId, msg || undefined, id);
      }
    },
    [ctx, id]
  );

  const validate = React.useCallback((): string => {
    if (requiredEffective && selectedTags.length === 0) {
      return "This field is required.";
    }
    return "";
  }, [requiredEffective, selectedTags]);

  // -------------------------------------------------------------------
  // Commit numeric SPUserID values into GlobalFormData
  // (mirrors TagPicker's commitValue pattern)
  // -------------------------------------------------------------------
  const commitValue = React.useCallback(() => {
    const err = validate();
    reportError(err);

    const targetId = `${id}Id`;

    // extract numeric IDs from lastResolved by matching Tag keys
    const ids: number[] = [];
    const keyed = new Map<string, PickerEntity>();
    for (const e of lastResolved) {
      const key =
        e.Key ??
        e.EntityData2?.AccountName ??
        e.EntityData2?.Email ??
        e.DisplayText ??
        "";
      keyed.set(String(key).toLowerCase(), e);
    }

    for (const t of selectedTags) {
      const lk = String(t.key).toLowerCase();
      const ent = keyed.get(lk);
      if (!ent) continue;

      // SPUserID comes back as Key when we hydrate from /getUserById
      const asNum = Number(ent.Key);
      if (!Number.isNaN(asNum)) {
        ids.push(asNum);
      }
    }

    if (ctx.GlobalFormData) {
      if (isMulti) {
        ctx.GlobalFormData(targetId, ids.length === 0 ? [] : ids);
      } else {
        ctx.GlobalFormData(
          targetId,
          ids.length === 0 ? null : ids[0]
        );
      }
    }

    // pretty label text used by display-only rendering
    const labels = selectedTags.map((t) => t.name);
    setDisplayOverride(labels.join("; "));

    if (ctx.GlobalRefs) {
      ctx.GlobalRefs(
        elemRef.current !== null ? elemRef.current : undefined
      );
    }

    // also push full entities out to parent if they provided onChange2
    if (onChange2) {
      onChange2(lastResolved.filter((e) =>
        selectedTags.some((t) =>
          String(t.key).toLowerCase() ===
          (e.Key ??
            e.EntityData2?.AccountName ??
            e.EntityData2?.Email ??
            e.DisplayText ??
            ""
          ).toString().toLowerCase()
        )
      ));
    }
  }, [
    id,
    isMulti,
    lastResolved,
    onChange2,
    reportError,
    selectedTags,
    validate,
    ctx
  ]);

  const handleBlur = React.useCallback(() => {
    setTouched(true);
    commitValue();
  }, [commitValue]);

  // -------------------------------------------------------------------
  // Search: REST people API (ClientPeoplePickerSearchUser)
  // -------------------------------------------------------------------
  const searchPeople = React.useCallback(
    async (query: string): Promise<ITag[]> => {
      if (!query.trim()) return [];

      const body = JSON.stringify({
        queryParams: {
          __metadata: {
            type:
              "SP.UI.ApplicationPages.ClientPeoplePickerQueryParameters"
          },
          AllowEmailAddresses: true,
          AllowMultipleEntities: isMulti,
          AllUrlZones: false,
          MaximumEntitySuggestions: maxSuggestions,
          QueryString: query,
          PrincipalSource: 1,
          PrincipalType: principalType
        }
      });

      try {
        // Prefer SPFx client when provided
        if (spHttpClient && spHttpClientConfig) {
          const resp = await spHttpClient.post(
            apiUrl,
            spHttpClientConfig,
            {
              headers: {
                Accept: "application/json;odata=verbose",
                "Content-Type": "application/json;odata=verbose",
                "odata-version": "3.0"
              },
              body
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

        // Fallback: plain fetch with request digest
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
            "odata-version": "3.0"
          },
          body,
          credentials: "same-origin"
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
        const raw = json.d?.ClientPeoplePickerSearchUser ?? "[]";
        const entities: PickerEntity[] = JSON.parse(raw);
        setLastResolved(entities);
        return entities.map(toTag);
      } catch (e) {
        console.error("PeoplePicker search exception:", e);
        return [];
      }
    },
    [
      apiUrl,
      isMulti,
      principalType,
      maxSuggestions,
      spHttpClient,
      spHttpClientConfig
    ]
  );

  // -------------------------------------------------------------------
  // TagPicker -> entities mapping when user changes selection
  // -------------------------------------------------------------------
  const handleChange = React.useCallback(
    (items?: ITag[]) => {
      const next = items ?? [];

      // For single-select we *also* rely on itemLimit={1},
      // but keeping this guard is harmless.
      const clamped = !isMulti && next.length > 1
        ? [next[next.length - 1]]
        : next;

      setSelectedTags(clamped);

      // We *don’t* commit here – we defer to blur/submit to keep
      // behavior consistent with other fields.
      if (!touched) return;

      commitValue();
    },
    [commitValue, isMulti, touched]
  );

  // -------------------------------------------------------------------
  // EDIT / VIEW FORM: hydrate from ctx.FormData (SPUserID values)
  // -------------------------------------------------------------------
  const hydratedFromForm = React.useRef(false);

  React.useEffect(() => {
    if (hydratedFromForm.current) return;
    if (!(ctx.FormMode === 4 || ctx.FormMode === 6)) return;

    const formData = ctx.FormData as any | undefined;
    if (!formData) return;

    const fldInternalName = id;
    let rawValue: any =
      formData[`${fldInternalName}Id`] ??
      formData[`${fldInternalName}IdString`] ??
      formData[fldInternalName];

    if (rawValue === null || rawValue === undefined) return;

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
                Accept: "application/json;odata=verbose"
              },
              signal: abort.signal
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
              Department2: u.Department || ""
            }
          });
        } catch (err) {
          if (abort.signal.aborted) return;
          console.error("PeoplePicker getUserById error", err);
        }
      }

      if (!hydrated.length) return;

      hydratedFromForm.current = true;
      setLastResolved(hydrated);
      const tags = hydrated.map(toTag);
      setSelectedTags(tags);

      // also show text + push entities up
      const labels = tags.map((t) => t.name);
      setDisplayOverride(labels.join("; "));
      if (onChange2) {
        onChange2(hydrated);
      }
    })();

    return () => abort.abort();
  }, [ctx.FormMode, ctx.FormData, id, onChange2, webUrl]);

  // -------------------------------------------------------------------
  // NEW FORM: start from starterValue and clear on mode change
  // -------------------------------------------------------------------
  React.useEffect(() => {
    if (ctx.FormMode === 8) {
      // New form
      if (normalizedStarter.length) {
        setSelectedTags(normalizedStarter);
      } else {
        setSelectedTags([]);
        setLastResolved([]);
      }
      setDisplayOverride("");
      setTouched(false);
      hydratedFromForm.current = false;
    }
  }, [ctx.FormMode, normalizedStarter]);

  // -------------------------------------------------------------------
  // Submitting disables and locks display text (same idea as TagPicker)
  // -------------------------------------------------------------------
  React.useEffect(() => {
    if (defaultDisable === false) {
      setIsDisabled(!!submitting || !!disabled);
    } else {
      setIsDisabled(true);
    }
  }, [submitting, disabled, defaultDisable]);

  // Keep GlobalRefs up-to-date when selection changes
  React.useEffect(() => {
    if (ctx.GlobalRefs) {
      ctx.GlobalRefs(
        elemRef.current !== null ? elemRef.current : undefined
      );
    }
  }, [ctx.GlobalRefs, selectedTags]);

  // -------------------------------------------------------------------
  // Rendering
  // -------------------------------------------------------------------
  const requiredMsg =
    requiredEffective && selectedTags.length === 0
      ? "This field is required."
      : undefined;

  const hasError = !!(errorMsg || requiredMsg);
  const validationMsg = errorMsg || requiredMsg;

  const disabledFinal = isDisabled;

  return (
    <div
      ref={elemRef}
      style={{ display: isHidden ? "none" : "block" }}
      className="fieldClass"
      aria-disabled={disabledFinal ? "true" : undefined}
      data-disabled={disabledFinal ? "true" : undefined}
    >
      <Field
        label={displayName}
        hint={description}
        validationMessage={validationMsg}
        validationState={hasError ? "error" : "none"}
      >
        {displayOverride && disabledFinal ? (
          // Display form: keep text visible but non-interactive
          <div style={{ marginTop: 4, fontSize: 12, opacity: 0.7 }}>
            {displayOverride}
          </div>
        ) : (
          <TagPicker
            className={className}
            disabled={disabledFinal}
            // Fluent UI single-select pattern
            itemLimit={isMulti ? undefined : 1}
            onResolveSuggestions={(filter, selected) => {
              const current = selected ?? [];
              // Extra guard: if single and one already selected, stop
              if (!isMulti && current.length >= 1) {
                return [];
              }
              if (!filter || !filter.trim()) return [];
              return searchPeople(filter).then((tags) =>
                tags.filter(
                  (t) =>
                    current.every(
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
              onBlur: handleBlur
            }}
          />
        )}
      </Field>
    </div>
  );
};

export default PeoplePicker;





