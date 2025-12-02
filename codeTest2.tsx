// PeoplePickerComponent.tsx

import * as React from "react";
import {
  Field,
  Textarea,
  useId
} from "@fluentui/react-components";
import {
  TagPicker,
  TagPickerControl,
  TagPickerGroup,
  TagPickerInput,
  TagPickerList,
  TagPickerOption,
  useTagPickerFilter,
  Tag
} from "@fluentui/react-components/unstable";
import { DynamicFormContext } from "./DynamicFormContext";

// -----------------------------------------------------------
// Types
// -----------------------------------------------------------

export interface PeoplePickerEntity {
  Id: number;
  Title: string;
  Email?: string;
  LoginName?: string;
}

export interface PeoplePickerOption {
  /** SPUserID – must be numeric, but we keep it as string for TagPicker */
  key: string;
  /** Display name */
  text: string;
  /** Optional email / extra text if you ever want to show it */
  secondaryText?: string;
}

export interface PeoplePickerProps {
  id: string;
  displayName?: string;
  className?: string;
  description?: string;
  placeholder?: string;

  isRequired?: boolean;
  isrequired2?: boolean; // kept for parity with other components
  submitting?: boolean;
  multiselect?: boolean;
  disabled?: boolean;

  /** Optional starter value(s) for NEW form (SPUserID + display text)  */
  starterValue?:
    | { key: number; text: string }
    | { key: number; text: string }[];

  /** Optional knobs for people search */
  principalType?: 0 | 1 | 2 | 4 | 8 | 15; // default 1 = user
  maxSuggestions?: number;

  /** Optional SPFx client – if supplied we’ll use POST with it */
  spHttpClient?: any;
  spHttpClientConfig?: any;
}

// -----------------------------------------------------------
// Constants / helpers
// -----------------------------------------------------------

const REQUIRED_MSG = "This field is required.";

/**
 * Convert a PeoplePickerOption key to its numeric SPUserID.
 * (All keys are stored as the numeric ID in string form.)
 */
const keyToUserId = (optionKey: string): number =>
  Number(optionKey);

/**
 * Normalize any SP “Id / Id.results / string of IDs” value into an array of
 * numeric SPUserIDs.
 */
const normalizeIdsFromFormData = (raw: any): number[] => {
  if (raw === null || raw === undefined) return [];

  // Multi-value lookup: { results: [id, id, …] }
  if (typeof raw === "object" && Array.isArray(raw.results)) {
    return raw.results
      .map((v: any) => Number(v))
      .filter((v: number) => !Number.isNaN(v));
  }

  // Plain array of IDs
  if (Array.isArray(raw)) {
    return raw
      .map((v: any) => Number(v))
      .filter((v: number) => !Number.isNaN(v));
  }

  // Single numeric ID (or numeric string)
  const n = Number(raw);
  return Number.isNaN(n) ? [] : [n];
};

/**
 * Build a map from option key -> display text, just like your TagPicker.
 */
const buildKeyToText = (options: PeoplePickerOption[]): Map<string, string> => {
  const m = new Map<string, string>();
  for (const o of options) {
    m.set(o.key, o.text);
  }
  return m;
};

// -----------------------------------------------------------
// Component
// -----------------------------------------------------------

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
    principalType = 1,
    maxSuggestions = 5,
    spHttpClient,
    spHttpClientConfig
  } = props;

  const requiredEffective = (isRequired ?? isrequired2) ?? false;
  const isMulti = multiselect === true;

  // Explicit site URL – match the pattern you used in TagPicker
  const webUrl =
    "https://amerihealthcaritas.sharepoint.com/sites/eokm";
  const apiSearchUrl =
    `${webUrl}/_api/SP.UI.ApplicationPages.ClientPeoplePickerWebServiceInterface.ClientPeoplePickerSearchUser`;
  const apiGetUserUrl = `${webUrl}/_api/web/getUserById`;

  // -------------------------------------------------------
  // Global UI state (disabled, hidden, required, etc.)
  // Mirrors TagPickerComponent.tsx
  // -------------------------------------------------------

  const [query, setQuery] = React.useState<string>("");
  const [options, setOptions] = React.useState<PeoplePickerOption[]>([]);
  const [selectedOptions, setSelectedOptions] = React.useState<string[]>([]);

  const [isDisabled, setIsDisabled] = React.useState<boolean>(
    !!disabled
  );
  const [defaultDisable, setDefaultDisable] = React.useState<boolean>(
    false
  );
  const [isHidden, setIsHidden] = React.useState<boolean>(false);
  const [touched, setTouched] = React.useState<boolean>(false);
  const [errorMsg, setErrorMsg] = React.useState<string>("");
  const [displayOverride, setDisplayOverride] = React.useState<string>("");

  const tagId = useId("default");

  // This ref is what GlobalRefs expects – a plain input element
  const elemRef = React.useRef<HTMLInputElement | null>(null);

  // -------------------------------------------------------
  // Error reporting (matches TagPickerComponent)
  // -------------------------------------------------------

  const reportError = React.useCallback(
    (msg: string) => {
      const targetId = `${id}Id`;
      setErrorMsg(msg || "");
      ctx.GlobalErrorHandle(targetId, msg || undefined);
    },
    [ctx, id]
  );

  // -------------------------------------------------------
  // People search (REST API)
  // -------------------------------------------------------

  const searchPeople = React.useCallback(
    async (text: string): Promise<PeoplePickerOption[]> => {
      const trimmed = text.trim();
      if (!trimmed) {
        return [];
      }

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
          PrincipalSource: 1,
          PrincipalType: principalType,
          QueryString: trimmed
        }
      });

      try {
        // Prefer SPFx client if provided
        if (spHttpClient && spHttpClientConfig) {
          const resp = await spHttpClient.post(
            apiSearchUrl,
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
              "PeoplePicker spHttpClient error",
              resp.status,
              resp.statusText,
              txt
            );
            return [];
          }

          const data: any = await resp.json();
          const raw =
            data.d?.ClientPeoplePickerSearchUser ?? "[]";
          const entities: any[] = JSON.parse(raw);

          const mapped: PeoplePickerOption[] = entities.map(
            (e: any) => {
              const userId =
                e.EntityData?.SPUserID ??
                e.EntityData?.UserId ??
                e.Key;
              return {
                key: String(userId),
                text: e.DisplayText ?? e.Key ?? "",
                secondaryText: e.EntityData?.Email
              };
            }
          );

          setOptions(mapped);
          return mapped;
        }

        // Fallback: classic fetch with request digest
        const digest = (document.getElementById(
          "__REQUESTDIGEST"
        ) as HTMLInputElement | null)?.value;

        const resp = await fetch(apiSearchUrl, {
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
            "PeoplePicker fetch error",
            resp.status,
            resp.statusText,
            txt
          );
          return [];
        }

        const json: any = await resp.json();
        const raw =
          json.d?.ClientPeoplePickerSearchUser ?? "[]";
        const entities: any[] = JSON.parse(raw);

        const mapped: PeoplePickerOption[] = entities.map(
          (e: any) => {
            const userId =
              e.EntityData?.SPUserID ??
              e.EntityData?.UserId ??
              e.Key;
            return {
              key: String(userId),
              text: e.DisplayText ?? e.Key ?? "",
              secondaryText: e.EntityData?.Email
            };
          }
        );

        setOptions(mapped);
        return mapped;
      } catch (e) {
        console.error("PeoplePicker search exception:", e);
        return [];
      }
    },
    [
      apiSearchUrl,
      isMulti,
      maxSuggestions,
      principalType,
      spHttpClient,
      spHttpClientConfig
    ]
  );

  // -------------------------------------------------------
  // Suggestions + TagPicker children (mirrors TagPickerComponent)
  // -------------------------------------------------------

  const optionTexts = React.useMemo(
    () => options.map((v) => v.text),
    [options]
  );

  const noMatchText = "We couldn't find any matches";

  const children = useTagPickerFilter({
    query,
    options: optionTexts,
    noOptionsElement: (
      <TagPickerOption value="no-matches">
        {noMatchText}
      </TagPickerOption>
    ),
    filter: (option) =>
      !selectedOptions.includes(option) &&
      option.toLowerCase().includes(query.toLowerCase())
  });

  const keyToText = React.useMemo(
    () => buildKeyToText(options),
    [options]
  );

  // -------------------------------------------------------
  // Validation + committing numeric IDs into GlobalFormData
  // -------------------------------------------------------

  const validate = React.useCallback(
    (): string => {
      if (!requiredEffective) return "";
      if (selectedOptions.length === 0) return REQUIRED_MSG;
      return "";
    },
    [requiredEffective, selectedOptions.length]
  );

  const commitValue = React.useCallback(() => {
    const err = validate();
    reportError(err);

    const targetId = `${id}Id`;
    const nums: number[] = [];

    for (let i = 0; i < selectedOptions.length; i++) {
      const label = selectedOptions[i];
      const opt = options.find((o) => o.text === label);
      if (opt) {
        nums.push(keyToUserId(opt.key));
      }
    }

    if (isMulti) {
      ctx.GlobalFormData(targetId, nums.length === 0 ? [] : nums);
    } else {
      ctx.GlobalFormData(targetId, nums.length === 0 ? null : nums[0]);
    }

    const labels = selectedOptions.map(
      (k) => keyToText.get(k) ?? k
    );
    setDisplayOverride(labels.join("; "));

    ctx.GlobalRefs(
      elemRef.current !== null ? elemRef.current : undefined,
      id
    );
  }, [
    ctx,
    id,
    isMulti,
    keyToText,
    options,
    reportError,
    selectedOptions,
    validate
  ]);

  const handleBlur = React.useCallback((): void => {
    setTouched(true);
    commitValue();
  }, [commitValue]);

  // -------------------------------------------------------
  // TagPicker selection logic – multi vs single select
  // -------------------------------------------------------

  const onOptionSelect: React.ComponentProps<
    typeof TagPicker
  >["onOptionSelect"] = (e, data) => {
    const next = (data.selectedOptions ?? []).map((v) => v);

    if (touched) {
      const err = validate();
      reportError(err);
    }

    if (data.value === "no-matches") {
      return;
    }

    if (isMulti) {
      setSelectedOptions(next);
    } else {
      // Single-select behaviour: keep only the latest value if present
      if (data.value !== undefined) {
        if (data.selectedOptions.length === 0) {
          const single = [data.value];
          setSelectedOptions(single);
        } else {
          setSelectedOptions([data.value]);
        }
      }
    }

    setQuery("");
  };

  // -------------------------------------------------------
  // Submitting: disable field + lock display text
  // -------------------------------------------------------

  React.useEffect(() => {
    if (defaultDisable === false && submitting) {
      setIsDisabled(true);
    } else {
      setIsDisabled(false);
      const labels = selectedOptions.map(
        (k) => keyToText.get(k) ?? k
      );
      setDisplayOverride(labels.join("; "));
    }

    const next = (selectedOptions ?? []).map((k) => k);
    if (touched) {
      const err =
        requiredEffective && next.length === 0
          ? REQUIRED_MSG
          : "";
      reportError(err);
    }
  }, [
    submitting,
    defaultDisable,
    selectedOptions,
    keyToText,
    touched,
    reportError,
    requiredEffective
  ]);

  // -------------------------------------------------------
  // Initial render + default / edit / view hydration
  // -------------------------------------------------------

  React.useEffect(() => {
    // Ensure the elemRef is registered in GlobalRefs (like TagPicker)
    ctx.GlobalRefs(
      elemRef.current !== null ? elemRef.current : undefined,
      id
    );

    // ---------- NEW FORM (FormMode === 8) ----------
    if (ctx.FormMode === 8) {
      if (starterValue === undefined) {
        setSelectedOptions([]);
        setOptions([]);
        setDisplayOverride("");
        setTouched(false);
        reportError("");
        return;
      }

      const starterArray = Array.isArray(starterValue)
        ? starterValue
        : [starterValue];

      const starterOptions: PeoplePickerOption[] =
        starterArray.map((s) => ({
          key: String(s.key),
          text: s.text
        }));

      setOptions(starterOptions);
      setSelectedOptions(starterOptions.map((s) => s.text));
      const labels = starterOptions.map((s) => s.text);
      setDisplayOverride(labels.join("; "));
      setTouched(false);
      reportError("");
      return;
    }

    // ---------- EDIT / VIEW FORM ----------
    const formData = ctx.FormData as any | undefined;
    if (!formData) return;

    const internalIdProp = `${id}Id`;
    const rawValue = formData[internalIdProp];

    const numericIds = normalizeIdsFromFormData(rawValue);
    if (!numericIds.length) {
      setSelectedOptions([]);
      setDisplayOverride("");
      setTouched(false);
      reportError("");
      return;
    }

    // Fetch each user by ID and hydrate options + selections
    const abort = new AbortController();

    (async () => {
      const hydratedOptions: PeoplePickerOption[] = [];

      for (const userId of numericIds) {
        try {
          const resp = await fetch(
            `${apiGetUserUrl}(${userId})`,
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

          hydratedOptions.push({
            key: String(u.Id),
            text: u.Title,
            secondaryText: u.Email
          });
        } catch (err) {
          if (abort.signal.aborted) return;
          console.error("PeoplePicker getUserById error", err);
        }
      }

      if (!hydratedOptions.length) {
        setSelectedOptions([]);
        setOptions([]);
        setDisplayOverride("");
        setTouched(false);
        reportError("");
        return;
      }

      setOptions(hydratedOptions);
      const labels = hydratedOptions.map((o) => o.text);
      setSelectedOptions(labels);
      setDisplayOverride(labels.join("; "));

      // For Display form, mark disabled + run formFieldsSetup, just like TagPicker
      if (ctx.FormMode === 4) {
        setIsDisabled(true);
      }

      setTouched(false);
      reportError("");
    })();

    return () => abort.abort();
  }, [
    ctx.FormMode,
    ctx.FormData,
    id,
    starterValue,
    apiGetUserUrl,
    ctx,
    reportError
  ]);

  // -------------------------------------------------------
  // Display text + click to remove tags
  // -------------------------------------------------------

  const joinedText = selectedOptions.join("; ");
  const visibleText = displayOverride || joinedText;
  const triggerText = visibleText || "";
  const triggerPlaceholder = triggerText || (placeholder || "");

  const rootClassName = [className, isDisabled ? "is-disabled" : ""]
    .filter(Boolean)
    .join(" ");

  const onTagClick = (option: string): void => {
    // Remove selected option
    const remainder = selectedOptions.filter((o) => o !== option);
    setSelectedOptions(remainder);

    const targetId = `${id}Id`;
    const nums: number[] = [];
    for (let i = 0; i < remainder.length; i++) {
      const label = remainder[i];
      const opt = options.find((t) => t.text === label);
      if (opt) {
        nums.push(keyToUserId(opt.key));
      }
    }

    if (isMulti) {
      ctx.GlobalFormData(targetId, nums.length === 0 ? [] : nums);
    } else {
      ctx.GlobalFormData(targetId, nums.length === 0 ? null : nums[0]);
    }

    const labels = remainder.map(
      (k) => keyToText.get(k) ?? k
    );
    setDisplayOverride(labels.join("; "));
    ctx.GlobalRefs(
      elemRef.current !== null ? elemRef.current : undefined,
      id
    );
  };

  const hasError = !!errorMsg;

  // -------------------------------------------------------
  // Render
  // -------------------------------------------------------

  return (
    <div
      style={{ display: isHidden ? "none" : "block" }}
      className="fieldClass"
      aria-disabled={isDisabled ? "true" : undefined}
      data-disabled={isDisabled ? "true" : undefined}
    >
      <Field
        label={displayName}
        id={tagId}
        {...(requiredEffective && { required: true })}
        validationMessage={hasError ? errorMsg : undefined}
        validationState={hasError ? "error" : undefined}
      >
        {isDisabled ? (
          // Disabled input to retain gray-out visuals and keep text visible
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
            selectionMode={isMulti ? "multiselect" : "single"}
            positioning="below-end"
          >
            <TagPickerControl>
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
                placeholder={placeholder}
                onChange={(e) => {
                  const value = e.target.value;
                  setQuery(value);
                  // Live search while typing
                  void searchPeople(value);
                }}
                onBlur={handleBlur}
              />
            </TagPickerControl>
            {/* Drop-down list */}
            <TagPickerList className="tagpickerList">
              {children}
            </TagPickerList>
          </TagPicker>
        )}

        {/* Hidden input: mirrors TagPicker – used by GlobalRefs & form submit */}
        <input
          style={{ display: "none" }}
          id={id}
          value={triggerText}
          ref={elemRef}
          readOnly
        />

        {description && (
          <div className="descriptionText">{description}</div>
        )}
      </Field>
    </div>
  );
};

export default PeoplePicker;







