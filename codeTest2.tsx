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
  Key: string;               // we’ll store SPUserID as string here
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

  /** starterValue: kept for compatibility with TagPicker */
  starterValue?:
    | { key: string | number; text: string }
    | { key: string | number; text: string }[];

  /** notify parent with SharePoint-style entities */
  onChange?: (entities: PickerEntity[]) => void;

  /* optional knobs (defaults supplied) */
  principalType?: PrincipalType;   // 1 = user only
  maxSuggestions?: number;         // default 5
  allowFreeText?: boolean;         // default false

  /* optional SPFx client + config for first-class POST */
  spHttpClient?: any;
  spHttpClientConfig?: any;

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
    onChange,
    principalType = 1,
    maxSuggestions = 5,
    allowFreeText = false,
    spHttpClient,
    spHttpClientConfig,
  } = props;

  const requiredEffective = (isRequired ?? isrequired2) ?? false;
  const isMulti = multiselect === true;

  // Explicit site URL – same pattern you used earlier
  const webUrl = "https://amerihealthcaritas.sharepoint.com/sites/eokm";
  const apiUrl =
    `${webUrl}/_api/SP.UI.ApplicationPages.ClientPeoplePickerWebServiceInterface.ClientPeoplePickerSearchUser`;

  /* ------------------------------------------------------------------ */
  /* Global UI state (mirrors other controls – disabled, hidden, etc.)  */
  /* ------------------------------------------------------------------ */

  const [isDisabled, setIsDisabled] = React.useState<boolean>(!!disabled);
  const [defaultDisable, setDefaultDisable] = React.useState<boolean>(false);
  const [isHidden, setIsHidden] = React.useState<boolean>(false);
  const [touched, setTouched] = React.useState<boolean>(false);
  const [errorMsg, setErrorMsg] = React.useState<string | undefined>();
  const [displayOverride, setDisplayOverride] = React.useState<string>("");

  // Ref used for GlobalRefs – put it on an outer <div>, not on TagPicker
  const elemRef = React.useRef<HTMLDivElement | null>(null);

  /* ------------------------------------------------------------------ */
  /* People-picker specific state                                      */
  /* ------------------------------------------------------------------ */

  // starterValue is only honored for NEW forms; Edit/View hydrating
  // will come from ctx.FormData.
  const starterArray: { key: string | number; text: string }[] =
    starterValue == null
      ? []
      : Array.isArray(starterValue)
      ? starterValue
      : [starterValue];

  const normalizedStarter: ITag[] = starterArray.map((v) => ({
    key: String(v.key),
    name: v.text,
  }));

  const [selectedTags, setSelectedTags] =
    React.useState<ITag[]>(normalizedStarter);
  const [lastResolved, setLastResolved] =
    React.useState<PickerEntity[]>([]);

  /* ------------------------------------------------------------------ */
  /* Error reporting via GlobalErrorHandle                              */
  /* ------------------------------------------------------------------ */

  const reportError = React.useCallback(
    (msg: string) => {
      const targetId = `${id}Id`;
      setErrorMsg(msg || "");
      ctx.GlobalErrorHandle(targetId, msg || undefined);
    },
    [ctx, id]
  );

  /* ------------------------------------------------------------------ */
  /* REST searchPeople (ClientPeoplePickerSearchUser)                   */
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
          PrincipalSource: 1,        // All
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
        const raw = json.d?.ClientPeoplePickerSearchUser ?? "[]";
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

  /* ------------------------------------------------------------------ */
  /* Commit current selection into GlobalFormData + onChange            */
  /* ------------------------------------------------------------------ */

  const commitSelection = React.useCallback(
    (items: ITag[]) => {
      const result: PickerEntity[] = [];
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

      for (const t of items) {
        const lk = String(t.key).toLowerCase();
        const match = resolvedByKey.get(lk);

        if (match) {
          result.push(match);
        } else if (allowFreeText) {
          // synthesize minimal entity from free-text if allowed
          result.push({
            Key: String(t.key),
            DisplayText: t.name,
            IsResolved: false,
            EntityType: "User",
            EntityData2: /@/.test(String(t.key))
              ? { Email: String(t.key) }
              : undefined,
          });
        }
      }

      // Push numeric SPUserID values into GlobalFormData
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

      // Notify parent with full entities
      if (onChange) {
        onChange(result);
      }
    },
    [allowFreeText, ctx, id, isMulti, lastResolved, onChange]
  );

  /* ------------------------------------------------------------------ */
  /* Handle TagPicker onChange (enforce single vs multi, validate)      */
  /* ------------------------------------------------------------------ */

  const handleChange = React.useCallback(
    (items?: ITag[]) => {
      let next = items ?? [];

      // Enforce single-select by trimming list down to the **last** choice
      if (!isMulti && next.length > 1) {
        const last = next[next.length - 1];
        next = last ? [last] : [];
      }

      setSelectedTags(next);
      commitSelection(next);

      if (!touched) return;

      if (requiredEffective && next.length === 0) {
        reportError("This field is required.");
      } else {
        reportError("");
      }
    },
    [commitSelection, isMulti, requiredEffective, reportError, touched]
  );

  const handleBlur = React.useCallback(() => {
    setTouched(true);

    if (requiredEffective && selectedTags.length === 0) {
      reportError("This field is required.");
    } else {
      reportError("");
    }
  }, [requiredEffective, reportError, selectedTags.length]);

  /* ------------------------------------------------------------------ */
  /* EDIT / VIEW FORM: hydrate from ctx.FormData (SPUserID)             */
  /* ------------------------------------------------------------------ */

  React.useEffect(() => {
    // Only run for Edit (6) or View (4)
    if (ctx.FormMode !== 4 && ctx.FormMode !== 6) return;

    const formData = ctx.FormData as any | undefined;
    if (!formData) return;

    const fieldInternalName = id;
    let rawValue: any = formData[fieldInternalName];

    // Handle the various SP shapes:
    //  - <InternalName>Id
    //  - <InternalName>StringId
    if (rawValue === undefined) {
      const idProp = `${fieldInternalName}Id`;
      const stringIdProp = `${fieldInternalName}StringId`;
      rawValue = formData[idProp] ?? formData[stringIdProp];
    }

    if (rawValue === null || rawValue === undefined) return;

    // Normalize into numeric SPUserID[] from whatever SP gave us
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

      // string like "1;#2;#3"
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
            Key: String(u.Id),            // SPUserID
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

      // Commit the hydrated value into GlobalFormData + fire onChange
      commitSelection(tags);
    })();

    return () => abort.abort();
  }, [ctx.FormMode, ctx.FormData, id, commitSelection, webUrl]);

  /* ------------------------------------------------------------------ */
  /* NEW FORM: ensure picker starts with a clean state                  */
  /* ------------------------------------------------------------------ */

  React.useEffect(() => {
    if (ctx.FormMode === 8) {
      // New form – if starterValue exists, we already seeded it in state;
      // otherwise start clean so search works normally.
      if (!starterArray.length) {
        setSelectedTags([]);
        setLastResolved([]);
      }
    }
  }, [ctx.FormMode, starterArray.length]);

  /* ------------------------------------------------------------------ */
  /* Disable / hide logic mirroring TextArea / TagPicker                */
  /* ------------------------------------------------------------------ */

  React.useEffect(() => {
    // For Display form, field is disabled
    if (ctx.FormMode === 4) {
      setIsDisabled(true);
    } else {
      const formFieldProps: any = {
        disabledList: ctx.AllDisableFields,
        hiddenList: ctx.AllHiddenFields,
        userBasedList: ctx.userBasedPerms,
        curUserList: ctx.curUserInfo,
        curField: id,
        formStateData: ctx.FormData,
        listColumns: ctx.listCols,
      };

      const results = ctx.formFieldsSetup
        ? ctx.formFieldsSetup(formFieldProps)
        : [];

      if (results && results.length > 0) {
        for (let i = 0; i < results.length; i++) {
          if (results[i].isDisabled !== undefined) {
            setDefaultDisable(results[i].isDisabled);
            setIsDisabled(results[i].isDisabled);
          }
          if (results[i].isHidden !== undefined) {
            setIsHidden(results[i].isHidden);
          }
        }
      }
    }

    // Register ref for scroll / focus behaviour
    ctx.GlobalRefs(
      elemRef.current !== null ? elemRef.current : undefined,
      id
    );
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []); // intentionally run once (mirrors TagPicker/TextArea)

  React.useEffect(() => {
    // When submitting, lock the field only if rules didn’t already disable it
    if (!defaultDisable) {
      setIsDisabled(!!submitting);
    }
  }, [defaultDisable, submitting]);

  /* ------------------------------------------------------------------ */
  /* Validation message for <Field>                                    */
  /* ------------------------------------------------------------------ */

  const requiredMsg =
    requiredEffective && selectedTags.length === 0 && touched
      ? "This field is required."
      : undefined;

  const validationMsg = requiredMsg ?? errorMsg;
  const hasError = !!validationMsg;

  /* ------------------------------------------------------------------ */
  /* Rendering                                                          */
  /* ------------------------------------------------------------------ */

  return (
    <div
      ref={elemRef}
      style={{ display: isHidden ? "none" : "block" }}
      className={className}
      data-disabled={isDisabled ? "true" : undefined}
    >
      <Field
        label={displayName}
        hint={description}
        validationMessage={validationMsg}
        validationState={hasError ? "error" : "none"}
        required={requiredEffective}
      >
        {displayOverride && (
          <div style={{ marginTop: 4, fontSize: 12, opacity: 0.7 }}>
            {displayOverride}
          </div>
        )}

        <TagPicker
          disabled={isDisabled}
          itemLimit={isMulti ? undefined : 1}
          onResolveSuggestions={(filter, selected) => {
            const already = selected ?? [];
            if (!filter.trim() && already.length === 0) return [];
            return searchPeople(filter).then((tags) =>
              tags.filter(
                (t) =>
                  !already.some((s) => String(s.key) === String(t.key))
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
          }}
        />
      </Field>
    </div>
  );
};

export default PeoplePicker;





