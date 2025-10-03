import * as React from 'react';
import {
  Field,
  Dropdown,
  Option,
  Input,
  type OptionOnSelectData,
  type SelectionEvents,
} from '@fluentui/react-components';
import { DynamicFormContext } from './DynamicFormContext';
import formFieldsSetup, { type FormFieldsProps } from './formFieldBased';

/* ---------- props, same shape as your original ---------- */
interface DropdownProps {
  id: string;
  starterValue?: string | number | Array<string | number>;
  displayName: string;
  isRequired?: boolean;
  placeholder?: string;
  multiSelect?: boolean;     // v8 prop
  multiselect?: boolean;     // v9 prop
  fieldType?: string;        // 'lookup' => commit under `${id}Id` as numbers
  options: { key: string | number; text: string }[];
  className?: string;
  description?: string;
  disabled?: boolean;
  submitting?: boolean;
}

/* ---------- constants & helpers (unchanged behavior) ---------- */
const REQUIRED_MSG = 'This is a required field and cannot be blank!';

// strict nullish -> string
const toKey = (k: unknown): string => (k === null || k === undefined ? '' : String(k));

function normalizeToStringArray(input: unknown): string[] {
  if (input === null || input === undefined) return [];

  // REST multi: { results: [] }
  if (
    typeof input === 'object' &&
    input !== null &&
    Array.isArray((input as { results?: unknown[] }).results)
  ) {
    return ((input as { results: unknown[] }).results).map(toKey);
  }

  if (Array.isArray(input)) {
    const arr = input as unknown[];
    if (arr.length > 0 && typeof arr[0] === 'object' && arr[0] !== null) {
      // Array of objects (e.g. [{ LookupId: 1 }, ...])
      return (arr as Array<Record<string, unknown>>).map(o =>
        toKey((o as { LookupId?: unknown; Id?: unknown }).LookupId ?? (o as { Id?: unknown }).Id ?? o)
      );
    }
    return (arr as Array<string | number>).map(toKey);
  }

  if (typeof input === 'string' && input.includes(';')) {
    return input.split(';').map(s => toKey(s.trim())).filter(Boolean);
  }

  return [toKey(input)];
}

function clampToExisting(values: string[], opts: { key: string | number }[]): string[] {
  const allowed = new Set(opts.map(o => toKey(o.key)));
  return values.filter(v => allowed.has(v));
}

/* =============================================================
   Component (structure & naming kept close to your screenshots)
   ============================================================= */
export default function DropdownComponent(props: DropdownProps): JSX.Element {
  // NOTE: Fluent v9 Dropdown trigger is a BUTTON → use HTMLButtonElement here.
  const elemRef = React.useRef<HTMLButtonElement | null>(null); // used to get the DOM element especially to get lookup values

  const {
    id,
    starterValue,
    displayName,
    isRequired: requiredProp = false,
    placeholder,
    multiSelect = false,
    fieldType,
    options,
    className,
    description,
    disabled: disabledProp = false,
    submitting = false,
  } = props;

  const isLookup = fieldType === 'lookup';

  // ---- context (same keys as in your screenshots) ----
  const {
    FormData,
    GlobalFormData,
    FormMode,
    GlobalErrorHandle,
    GlobalRefs,
    AllDisableFields,
    AllHiddenFields,
    userBasedPerms,
    curUserInfo,
    listCols,
  } = React.useContext(DynamicFormContext);

  // ---- local state (same semantics as original) ----
  const [isRequired, setIsRequired] = React.useState<boolean>(!!requiredProp);
  const [isDisabled, setIsDisabled] = React.useState<boolean>(!!disabledProp);
  const [isHidden, setIsHidden] = React.useState<boolean>(false);
  const [defaultIsDisable, setDefaultIsDisable] = React.useState<boolean>(false);

  const [selectedOptions, setSelectedOptions] = React.useState<string[]>([]);
  const [error, setError] = React.useState<string>('');
  const [touched, setTouched] = React.useState<boolean>(false);

  const [displayOverride, setDisplayOverride] = React.useState<string>('');
  const isLockedRef = React.useRef<boolean>(false);
  const didInitRef = React.useRef<boolean>(false);

  // map of key->text for labels
  const keyToText = React.useMemo<Map<string, string>>(() => {
    const m = new Map<string, string>();
    for (const o of options) m.set(toKey(o.key), o.text);
    return m;
  }, [options]);

  // error reporter (unchanged behavior)
  const reportError = React.useCallback((msg: string): void => {
    const targetId = isLookup ? `${id}LookupId` : id;
    setError(msg || '');
    GlobalErrorHandle?.(targetId, msg || null);
  }, [GlobalErrorHandle, id, isLookup]);

  // reflect prop changes to state
  React.useEffect((): void => {
    setIsRequired(!!requiredProp);
    setIsDisabled(!!disabledProp);
  }, [requiredProp, disabledProp]);

  // register element with your GlobalRefs (your context showed it as a function)
  React.useEffect((): void => {
    if (typeof GlobalRefs === 'function') {
      (GlobalRefs as (elmIntrnName?: string) => void)(id);
    }
  }, [GlobalRefs, id]);

  /* ------------------------------------------------------------------
     EFFECT 1: Init + sync from external data (condensed, dropdown-only)
     ------------------------------------------------------------------ */
  React.useEffect((): void => {
    if (isLockedRef.current) return;

    const ensureInOptions = (vals: string[]): string[] => clampToExisting(vals, options);
    const base = id.replace(/(Lookup)?Id$/i, '');

    if (!didInitRef.current) {
      // initialize from starterValue on create (FormMode != 3)
      if (FormMode !== 3) {
        const initArr = Array.isArray(starterValue)
          ? (starterValue as Array<string | number>).map(toKey)
          : [toKey(starterValue)];
        setSelectedOptions(ensureInOptions(initArr));
      }
      didInitRef.current = true;
      return;
    }

    // reflect current FormData -> control (only the cells this field cares about)
    const src =
      isLookup
        ? (multiSelect
            ? (FormData as Record<string, unknown> | undefined)?.[id] // array of {LookupId}
            : (FormData as Record<string, unknown> | undefined)?.[`${base}LookupId`])
        : (FormData as Record<string, unknown> | undefined)?.[id];

    const next = ensureInOptions(normalizeToStringArray(src));
    setSelectedOptions(next);

    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [
    FormMode,
    starterValue,
    options,
    id,
    isLookup,
    multiSelect,
    // specific FormData cells we read → avoids re-running on unrelated changes
    (FormData as any)?.[id],
    (FormData as any)?.[`${id.replace(/(Lookup)?Id$/i, '')}LookupId`],
  ]);

  /* --------------------------------------------------------
     EFFECT 2: Lock/disable when submitting or read-only mode
     -------------------------------------------------------- */
  React.useEffect((): void => {
    if (submitting || FormMode === 4) {
      setIsDisabled(true);
      isLockedRef.current = true;
      // compute labels once; no need to depend on selectedOptions here
      setDisplayOverride(selectedOptions.map(k => keyToText.get(k) ?? k).join('; '));
    }
  }, [submitting, FormMode, keyToText, selectedOptions]);

  /* ---------------------------------------------------------
     EFFECT 3: Rule-based disable/hide (uses your formFieldsSetup)
     --------------------------------------------------------- */
  React.useEffect((): void => {
    if (FormMode === 4) return;

    const formStateKeys: string[] = Array.isArray(FormData)
      ? (FormData as string[])
      : Object.keys((FormData ?? {}) as Record<string, unknown>);

    const listColumnsArray: string[] = Array.isArray(listCols)
      ? (listCols as string[]).map(String)
      : [];

    const formFieldProps: FormFieldsProps = {
      disabledList:  ((AllDisableFields ?? {}) as unknown as Record<string, unknown>),
      hiddenList:    ((AllHiddenFields ?? {}) as unknown as Record<string, unknown>),
      userBasedList: ((userBasedPerms ?? {}) as unknown as Record<string, unknown>),
      curUserList:   ((curUserInfo ?? {}) as unknown as Record<string, unknown>),
      curField:      displayName,
      formStateData: formStateKeys,
      listColumns:   listColumnsArray,
    };

    const results =
      (formFieldsSetup(formFieldProps) as Array<{ isDisabled?: boolean; isHidden?: boolean }>) ?? [];

    if (results.length) {
      const d = results[0]?.isDisabled;
      if (d !== undefined) {
        setIsDisabled(!!d);
        setDefaultIsDisable(!!d);
      }
      const h = results[0]?.isHidden;
      if (h !== undefined) setIsHidden(!!h);
    }
  }, [
    // minimal rule inputs only (kept close to your original names)
    AllDisableFields,
    AllHiddenFields,
    userBasedPerms,
    curUserInfo,
    listCols,
    displayName,
    FormMode,
    // re-run if the overall form-state shape changes
    Array.isArray(FormData) ? FormData.length : Object.keys(FormData ?? {}).length,
  ]);

  /* ---------- validation, commit, handlers (unchanged behavior) ---------- */
  const validate = React.useCallback((): string => {
    if (isRequired && selectedOptions.length === 0) return REQUIRED_MSG;
    return '';
  }, [isRequired, selectedOptions]);

  // Commit under `${id}Id` for lookup (numbers, null when empty). Non-lookup commits raw or first.
  const commitValue = React.useCallback((): void => {
    const err = validate();
    reportError(err);

    const targetId = isLookup ? `${id}Id` : id;

    if (isLookup) {
      const nums = selectedOptions
        .map(k => Number(k))
        .filter((n): n is number => Number.isFinite(n));
      if (multiSelect) {
        (GlobalFormData as (name: string, value: unknown) => void)(
          targetId,
          nums.length === 0 ? null : nums
        );
      } else {
        (GlobalFormData as (name: string, value: unknown) => void)(
          targetId,
          nums.length === 0 ? null : nums[0]
        );
      }
    } else {
      (GlobalFormData as (name: string, value: unknown) => void)(
        targetId,
        selectedOptions.length === 0 ? null : (multiSelect ? selectedOptions : selectedOptions[0])
      );
    }

    const labels = selectedOptions.map(k => keyToText.get(k) ?? k);
    setDisplayOverride(labels.join('; '));
  }, [validate, reportError, GlobalFormData, id, isLookup, multiSelect, selectedOptions, keyToText]);

  // Proper v9 types for handler
  const handleOptionSelect = React.useCallback(
    (_e: SelectionEvents, data: OptionOnSelectData): void => {
      const next = (data.selectedOptions ?? []).map(toKey);
      setSelectedOptions(next);
      if (!touched) reportError(isRequired && next.length === 0 ? REQUIRED_MSG : '');
    },
    [isRequired, reportError, touched]
  );

  const handleBlur = React.useCallback((): void => {
    setTouched(true);
    commitValue();
  }, [commitValue]);

  // joined labels for display
  const selectedLabels = selectedOptions.map(k => keyToText.get(k) ?? k);
  const joinedText = selectedLabels.join('; ');
  const triggerText = displayOverride || joinedText;
  const triggerPlaceholder = triggerText || (placeholder || '');
  const hasError = !!error;

  const disabledClass = isDisabled ? 'is-disabled' : '';
  const rootClassName = [className, disabledClass].filter(Boolean).join(' ');

  return (
    <div style={{ display: isHidden ? 'none' : 'block' }}>
      <Field
        label={displayName}
        required={isRequired}
        validationMessage={hasError ? error : undefined}
        validationState={hasError ? 'error' : undefined}
      >
        {isDisabled ? (
          // Disabled Input to retain gray-out visuals and keep text visible
          <Input
            id={id}
            disabled
            value={triggerText}
            placeholder={triggerPlaceholder}
            className={rootClassName}
            aria-disabled="true"
            data-disabled="true"
          />
        ) : (
          <Dropdown
            id={id}
            multiselect={!!multiSelect || !!multiselect}
            disabled={false}
            inlinePopup={true}
            selectedOptions={selectedOptions}
            onOptionSelect={handleOptionSelect}
            onBlur={handleBlur}
            className={rootClassName}
            value={triggerText}
            placeholder={triggerPlaceholder}
            title={triggerText || displayName}
            aria-label={triggerText || displayName}
            ref={elemRef}
          >
            {options.map(o => (
              <Option key={toKey(o.key)} value={toKey(o.key)}>
                {o.text}
              </Option>
            ))}
          </Dropdown>
        )}
      </Field>

      {description && <div className="descriptionText">{description}</div>}
    </div>
  );
}


