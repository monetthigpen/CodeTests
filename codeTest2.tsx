// PeoplePickerComponent.tsx
import * as React from "react";
import {
  TagPicker,
  TagPickerList,
  TagPickerInput,
  TagPickerControl,
  TagPickerGroup,
  TagPickerOption,
  TagPickerProps,
  useTagPickerFilter,
  Tag,
  Field,
  Textarea,
} from "@fluentui/react-components";

import { DynamicFormContext } from "./DynamicFormContext";
import { formFieldsSetup, FormFieldsProps } from "../utils/formFieldBased";

// ---------- Types ----------

export type PrincipalType = 0 | 1 | 2 | 4 | 8 | 15;

export interface PickerEntity {
  Key: string; // SharePoint user Id as string
  DisplayText?: string;
  EntityType?: string;
  IsResolved?: boolean;
  EntityData?: {
    Email?: string;
    AccountName?: string;
    Title?: string;
    Department?: string;
  };
}

export interface PeoplePickerProps {
  id: string;

  displayName?: string;
  className?: string;
  description?: string;
  placeholder?: string;

  isRequired?: boolean;
  isRequired2?: boolean;
  submitting?: boolean;

  multiselect?: boolean;
  disabled?: boolean;

  // Starter value: numeric user Id(s) – exactly what SP stores in the *Id field
  // e.g. { key: 50, text: 'Ada Lovelace' } or array of those
  starterValue?: { key: string | number; text: string } | Array<{
    key: string | number;
    text: string;
  }>;

  // People picker knobs
  principalType?: PrincipalType; // default 1 (User)
  maxSuggestions?: number; // default 5

  // Optional SPFx HTTP client – if not provided, falls back to classic fetch+digest
  spHttpClient?: any;
  spHttpClientConfig?: any;
}

// ---------- Constants & helpers ----------

const REQUIRED_MSG = "This is a required field and cannot be blank!";

const normalizeStarterArray = (
  starter?: PeoplePickerProps["starterValue"]
): Array<{ key: string; text: string }> => {
  if (!starter) return [];
  if (Array.isArray(starter)) {
    return starter.map((s) => ({
      key: String(s.key),
      text: s.text,
    }));
  }
  return [
    {
      key: String(starter.key),
      text: starter.text,
    },
  ];
};

// convert a PickerEntity into a simple display label
const entityToLabel = (e: PickerEntity): string => {
  return (
    e.DisplayText ||
    e.EntityData?.Title ||
    e.EntityData?.Email ||
    e.EntityData?.AccountName ||
    e.Key
  );
};

// collect numeric Ids (SPUserId) from SP form data (array or delimited string)
const collectUserIdsFromRaw = (rawValue: any): number[] => {
  if (rawValue == null) return [];

  if (Array.isArray(rawValue)) {
    return rawValue
      .map((v) => Number(v))
      .filter((id) => !Number.isNaN(id) && id > 0);
  }

  const str = String(rawValue);
  return str
    .split(/[;,#]/)
    .map((p) => Number(p.trim()))
    .filter((id) => !Number.isNaN(id) && id > 0);
};

// ---------- Component ----------

const PeoplePicker: React.FC<PeoplePickerProps> = (props) => {
  const ctx = React.useContext(DynamicFormContext);

  const {
    id,
    displayName,
    className,
    description,
    placeholder,
    isRequired,
    isRequired2,
    submitting,
    multiselect,
    disabled,
    principalType = 1,
    maxSuggestions = 5,
    spHttpClient,
    spHttpClientConfig,
  } = props;

  const isMulti = multiselect === true;
  const requiredEffective = (isRequired ?? isRequired2) ?? false;

  const webUrl =
    "https://amerihealthcaritas.sharepoint.com/sites/eokm"; // <-- adjust if needed

  // UI state – mirrors TagPickerComponent
  const [query, setQuery] = React.useState<string>("");
  const [selectedOptions, setSelectedOptions] = React.useState<string[]>([]);
  const [isDisabled, setIsDisabled] = React.useState<boolean>(!!disabled);
  const [defaultDisable, setDefaultDisable] = React.useState<boolean>(false);
  const [isHidden, setIsHidden] = React.useState<boolean>(false);
  const [touched, setTouched] = React.useState<boolean>(false);
  const [errorMsg, setErrorMsg] = React.useState<string>("");
  const [displayOverride, setDisplayOverride] = React.useState<string>("");

  // Suggestions from the PeoplePicker API
  const [options, setOptions] = React.useState<string[]>([]);

  // Last resolved entities (from search or hydration) – used for Id mapping
  const [lastResolved, setLastResolved] = React.useState<PickerEntity[]>([]);

  // ref to hidden input – used by GlobalRefs & for submission compatibility
  const elemRef = React.useRef<HTMLInputElement | null>(null);

  // ---------- Validation / Global error handling ----------

  const reportError = React.useCallback(
    (msg: string) => {
      const targetId = `${id}Id`;
      setErrorMsg(msg || "");
      ctx.GlobalErrorHandle(targetId, msg || undefined);
    },
    [ctx, id]
  );

  const validate = React.useCallback((): string => {
    if (!requiredEffective) return "";
    if (selectedOptions.length === 0) return REQUIRED_MSG;
    return "";
  }, [requiredEffective, selectedOptions.length]);

  // ---------- Utilities for mapping names <-> entities / Ids ----------

  const resolvedByLabel = React.useMemo(() => {
    const map = new Map<string, PickerEntity>();
    for (const e of lastResolved) {
      const label = entityToLabel(e);
      if (label) {
        map.set(label.toLowerCase(), e);
      }
    }
    return map;
  }, [lastResolved]);

  const getUserIdsFromSelection = React.useCallback((): number[] => {
    const ids: number[] = [];
    for (const label of selectedOptions) {
      const e = resolvedByLabel.get(label.toLowerCase());
      if (!e) continue;
      const num = Number(e.Key);
      if (!Number.isNaN(num) && num > 0) {
        ids.push(num);
      }
    }
    return ids;
  }, [selectedOptions, resolvedByLabel]);

  // ---------- Search (PeoplePicker Web Service) ----------

  const searchPeople = React.useCallback(
    async (queryText: string): Promise<string[]> => {
      const trimmed = queryText.trim();
      if (!trimmed) {
        return [];
      }

      const apiUrl = `${webUrl}/_api/SP.UI.ApplicationPages.ClientPeoplePickerWebServiceInterface.clientPeoplePickerSearchUser`;

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
          PrincipalSource: 1,
          PrincipalType: principalType,
          QueryString: trimmed,
        },
      });

      try {
        let resp: Response;

        if (spHttpClient && spHttpClientConfig) {
          resp = await spHttpClient.post(apiUrl, spHttpClientConfig, {
            headers: {
              Accept: "application/json;odata=verbose",
              "Content-Type": "application/json;odata=verbose",
              "odata-version": "3.0",
            },
            body,
          });
        } else {
          const digest =
            (document.getElementById(
              "__REQUESTDIGEST"
            ) as HTMLInputElement | null)?.value || "";

          resp = await fetch(apiUrl, {
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
        }

        if (!resp.ok) {
          const txt = await resp.text().catch(() => "");
          console.error(
            "PeoplePicker search error",
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

        const labels = entities
          .map(entityToLabel)
          .filter((s) => !!s) as string[];

        setOptions(labels);
        return labels;
      } catch (e) {
        console.error("PeoplePicker search exception", e);
        return [];
      }
    },
    [
      webUrl,
      isMulti,
      maxSuggestions,
      principalType,
      spHttpClient,
      spHttpClientConfig,
    ]
  );

  // ---------- TagPicker filter children (same pattern as TagPickerComponent) ----------

  const noMatchText = "We couldn't find any matches";

  const children = useTagPickerFilter({
    query,
    options,
    noOptionsElement: (
      <TagPickerOption value="no-matches">{noMatchText}</TagPickerOption>
    ),
    filter: (option: string) =>
      !selectedOptions.includes(option) &&
      option.toLowerCase().includes(query.toLowerCase()),
  });

  // ---------- Commit value to GlobalFormData ----------

  const commitValue = React.useCallback(() => {
    const err = validate();
    reportError(err);

    const targetId = `${id}Id`;
    const userIds = getUserIdsFromSelection();

    if (isMulti) {
      ctx.GlobalFormData(targetId, userIds.length === 0 ? [] : userIds);
    } else {
      ctx.GlobalFormData(
        targetId,
        userIds.length === 0 ? null : userIds[0]
      );
    }

    const labels = selectedOptions;
    setDisplayOverride(labels.join("; "));
    ctx.GlobalRefs(
      elemRef.current !== null ? elemRef.current : undefined,
      id
    );
  }, [
    ctx,
    id,
    isMulti,
    selectedOptions,
    getUserIdsFromSelection,
    reportError,
    validate,
  ]);

  // ---------- TagPicker event handlers ----------

  const onOptionSelect: TagPickerProps["onOptionSelect"] = React.useCallback(
    (_ev: React.SyntheticEvent, data: any) => {
      const value = data?.value as string | undefined;
      const nextSelected: string[] = data?.selectedOptions ?? [];

      if (touched) {
        const err = validate();
        reportError(err);
      }

      if (value === "no-matches") {
        return;
      }

      if (isMulti) {
        // multi: TagPicker already tracks all selected options for us
        setSelectedOptions(nextSelected);
      } else {
        // single: enforce one option only
        if (value != null && value !== "") {
          setSelectedOptions([value]);
        } else if (nextSelected.length > 0) {
          setSelectedOptions([nextSelected[0]]);
        } else {
          setSelectedOptions([]);
        }
      }

      setQuery("");
    },
    [isMulti, reportError, touched, validate]
  );

  const handleInputChange = React.useCallback(
    async (ev: React.ChangeEvent<HTMLInputElement>) => {
      const val = ev.target.value;
      setQuery(val);
      await searchPeople(val);
    },
    [searchPeople]
  );

  const handleBlur = React.useCallback((): void => {
    setTouched(true);
    commitValue();
  }, [commitValue]);

  // ---------- Submitting: disable & lock display text (same pattern as TagPicker) ----------

  React.useEffect(() => {
    if (!submitting && !defaultDisable) {
      // Form no longer submitting – re-enable if not default-disabled
      setIsDisabled(false);
      return;
    }

    if (submitting) {
      setIsDisabled(true);
      const labels = selectedOptions;
      setDisplayOverride(labels.join("; "));

      const next = selectedOptions;
      const err =
        requiredEffective && next.length === 0 ? REQUIRED_MSG : "";
      reportError(err);
    }
  }, [submitting, defaultDisable, selectedOptions, reportError, requiredEffective]);

  // ---------- Initial render / defaults / Edit / View hydration ----------

  React.useEffect(() => {
    ctx.GlobalRefs(
      elemRef.current !== null ? elemRef.current : undefined,
      id
    );

    // NEW FORM (8): use starterValue only
    if (ctx.FormMode === 8) {
      const starterArr = normalizeStarterArray(props.starterValue);
      const labels = starterArr.map((s) => s.text);
      setSelectedOptions(labels);
      setLastResolved(
        starterArr.map((s) => ({
          Key: String(s.key),
          DisplayText: s.text,
          IsResolved: true,
          EntityType: "User",
        }))
      );
      return;
    }

    // EDIT (6) / VIEW (4): hydrate from SPUserId values in ctx.FormData
    if (ctx.FormMode !== 4 && ctx.FormMode !== 6) {
      return;
    }

    const formData: any = ctx.FormData;
    if (!formData) return;

    const fieldInternalName = id;

    const idProp = `${fieldInternalName}Id`;
    const stringIdProp = `${fieldInternalName}IdStringId`;

    let rawValue = formData[idProp];
    if (rawValue === undefined || rawValue === null) {
      rawValue = formData[stringIdProp];
    }
    if (rawValue === undefined || rawValue === null) {
      return;
    }

    const numericIds = collectUserIdsFromRaw(rawValue);
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
            EntityData: {
              Email: u.Email,
              AccountName: u.LoginName,
              Title: u.Title,
              Department: u.Department || "",
            },
          });
        } catch (err) {
          if (abort.signal.aborted) return;
          console.error("PeoplePicker getUserById error", err);
        }
      }

      if (!hydrated.length) return;

      setLastResolved(hydrated);
      const labels = hydrated.map(entityToLabel);
      setSelectedOptions(labels);
    })();

    return () => abort.abort();
  }, [ctx.FormMode, ctx.FormData, ctx.GlobalRefs, id, props.starterValue, webUrl]);

  // ---------- Disable / hidden logic (same as TagPicker) ----------

  React.useEffect(() => {
    // Display form: always disabled, just show text
    if (ctx.FormMode === 4) {
      setIsDisabled(true);
      const labels = selectedOptions;
      setDisplayOverride(labels.join("; "));
      reportError("");
      setTouched(false);
      return;
    }

    // Edit / New: consult formFieldsSetup to see if this field is disabled/hidden
    const formFieldProps: FormFieldsProps = {
      disabledList: ctx.AllDisableFields,
      hiddenList: ctx.AllHiddenFields,
      userBasedList: ctx.userBasedPerms,
      curUserList: ctx.curUserInfo,
      curField: id,
      formStateData: ctx.FormData,
      listColumns: ctx.listCols,
    };

    const results = formFieldsSetup(formFieldProps);
    if (results.length > 0) {
      const r = results[0];
      if (r.isDisabled !== undefined) {
        setIsDisabled(r.isDisabled);
        setDefaultDisable(r.isDisabled);
      }
      if (r.isHidden !== undefined) {
        setIsHidden(r.isHidden);
      }
    }

    if (isDisabled) {
      const labels = selectedOptions;
      setDisplayOverride(labels.join("; "));
    }

    reportError("");
    setTouched(false);
  }, [
    ctx.FormMode,
    ctx.AllDisableFields,
    ctx.AllHiddenFields,
    ctx.userBasedPerms,
    ctx.curUserInfo,
    ctx.FormData,
    ctx.listCols,
    id,
    isDisabled,
    selectedOptions,
    reportError,
  ]);

  // ---------- Derived view values ----------

  const selectedLabels = selectedOptions;
  const joinedText = selectedLabels.join("; ");
  const visibleText = displayOverride || joinedText;
  const triggerText = visibleText || "";
  const triggerPlaceholder = triggerText || (placeholder || "");

  const hasError = !!errorMsg;
  const disabledClass = isDisabled ? "is-disabled" : "";
  const rootClassName = [className, disabledClass].filter(Boolean).join(" ");

  // remove one selected option (used when clicking an existing tag)
  const onTagClick = React.useCallback(
    (option: string): void => {
      const remainderOpts = selectedOptions.filter((o) => o !== option);
      setSelectedOptions(remainderOpts);

      const targetId = `${id}Id`;
      const userIds =
        remainderOpts.length === 0
          ? []
          : (() => {
              const ids: number[] = [];
              for (const label of remainderOpts) {
                const e = resolvedByLabel.get(label.toLowerCase());
                if (!e) continue;
                const num = Number(e.Key);
                if (!Number.isNaN(num) && num > 0) ids.push(num);
              }
              return ids;
            })();

      if (isMulti) {
        ctx.GlobalFormData(targetId, userIds.length === 0 ? [] : userIds);
      } else {
        ctx.GlobalFormData(
          targetId,
          userIds.length === 0 ? null : userIds[0]
        );
      }

      const labels = remainderOpts;
      setDisplayOverride(labels.join("; "));
      ctx.GlobalRefs(
        elemRef.current !== null ? elemRef.current : undefined,
        id
      );
    },
    [ctx, id, isMulti, selectedOptions, resolvedByLabel]
  );

  // ---------- Render ----------

  return (
    <div
      style={{ display: isHidden ? "none" : "block" }}
      className="fieldClass"
      aria-disabled={isDisabled ? "true" : undefined}
      data-disabled={isDisabled ? "true" : undefined}
    >
      <Field
        label={displayName}
        id={id}
        {...(requiredEffective && { required: true })}
        validationMessage={hasError ? errorMsg : undefined}
        validationState={hasError ? "error" : undefined}
      >
        {isDisabled ? (
          // Disabled Input to retain gray-out visuals and keep text visible
          <Textarea
            id={id}
            disabled
            value={triggerText}
            placeholder={triggerPlaceholder}
            className={rootClassName}
            aria-disabled="true"
            data-disabled="true"
          />
        ) : (
          <TagPicker
            size="medium"
            onOptionSelect={onOptionSelect}
            selectedOptions={selectedOptions}
            inline={true}
            positioning="below-end"
          >
            <TagPickerControl aria-label={displayName}>
              <TagPickerGroup aria-label={displayName}>
                {selectedOptions.map((option) => (
                  <Tag
                    key={option}
                    shape="rounded"
                    value={option}
                    className="lookupTags"
                    onClick={() => onTagClick(option)}
                  >
                    {option}
                  </Tag>
                ))}
              </TagPickerGroup>

              <TagPickerInput
                aria-label={displayName}
                value={query}
                onChange={handleInputChange}
                onBlur={handleBlur}
              />
            </TagPickerControl>

            {/* tagpickerList class is used to add z-index to drop down list */}
            <TagPickerList className="tagpickerList">
              {children}
            </TagPickerList>
          </TagPicker>
        )}

        {/* Hidden input field so that all selected options are added to an element
            which can be used later to get the text values for submission */}
        <input
          style={{ display: "none" }}
          id={id}
          value={triggerText}
          ref={elemRef}
          readOnly
        />
      </Field>

      {description && (
        <div className="descriptionText">{description}</div>
      )}
    </div>
  );
};

export default PeoplePicker;







