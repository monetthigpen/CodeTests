import * as React from 'react';
import { Field, Dropdown, Option, Input, type OptionOnSelectData, type SelectionEvents } from '@fluentui/react-components';
import { DynamicFormContext } from './DynamicFormContext';
import formFieldsSetup, { FormFieldsProps } from './formFieldBased';

interface DropdownProps {
  id: string;
  starterValue?: string | number | Array<string | number>;
  displayName: string;
  isRequired?: boolean;
  placeholder?: string;
  multiSelect?: boolean;    // v8 prop
  multiselect?: boolean;    // v9 prop
  fieldType?: string;       // 'lookup'  => commit under `${id}Id` as numbers
  options: { key: string | number; text: string }[];
  className?: string;
  description?: string;
  disabled?: boolean;
  submitting?: boolean;
}

const REQUIRED_MSG = 'This is a required field and cannot be blank!';

// strict (eqeqeq): catch both null and undefined
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
      // Array of objects (e.g., [{ LookupId: 1 }, ...])
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

export default function DropdownComponent(props: DropdownProps): JSX.Element {
  // NOTE: Fluent v9 Dropdown triggers a BUTTON -> ref must be HTMLButtonElement
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

  const {
    FormData,
    GlobalFormData,
    FormMode,
    GlobalErrorHandle,
    GlobalRefs,           // <- function in your context
    AllDisableFields,
    AllHiddenFields,
    userBasedPerms,
    curUserInfo,
    listCols,
  } = React.useContext(DynamicFormContext);

  const [isRequired, setIsRequired] = React.useState<boolean>(!!requiredProp);
  const [isDisabled, setIsDisabled] = React.useState<boolean>(!!disabledProp);
  const [isHidden, setIsHidden] = React.useState<boolean>(false);
  const [defaultIsDisable, setDefaultIsDisable] = React.useState<boolean>(false);

  // Controlled selection
  const [selectedOptions, setSelectedOptions] = React.useState<string[]>([]);
  const [error, setError] = React.useState<string>('');
  const [touched, setTouched] = React.useState<boolean>(false);

  // Lock/cache for display text when disabled
  const [displayOverride, setDisplayOverride] = React.useState<string>('');
  const isLockedRef = React.useRef<boolean>(false);
  const didInitRef = React.useRef<boolean>(false);

  const keyToText = React.useMemo<Map<string, string>>(() => {
    const m = new Map<string, string>();
    for (const o of options) m.set(toKey(o.key), o.text);
    return m;
  }, [options]);

  const reportError = React.useCallback((msg: string): void => {
    const targetId = isLookup ? `${id}LookupId` : id;
    setError(msg || '');
    GlobalErrorHandle?.(targetId, msg || null);
  }, [GlobalErrorHandle, id, isLookup]);

  React.useEffect((): void => {
    setIsRequired(!!requiredProp);
    setIsDisabled(!!disabledProp);
  }, [requiredProp, disabledProp]);

  // Submitting disables and locks display text
  React.useEffect((): void => {
    if (defaultIsDisable === false) {
      setIsDisabled(!!submitting);
    } else {
      setIsDisabled(true);
      isLockedRef.current = true;
      const labels = selectedOptions.map(k => keyToText.get(k) ?? k);
      setDisplayOverride(labels.join('; '));
    }
  }, [submitting, defaultIsDisable, selectedOptions, keyToText]);

  // Register element with your GlobalRefs FUNCTION (fixes the "Record<string, unknown>" error)
  React.useEffect((): void => {
    if (typeof GlobalRefs === 'function') {
      // your context shows GlobalRefs signature like (elmIntrnName: string | undefined) => void
      // call once with id (many apps just register key; pass element if your function accepts it)
      (GlobalRefs as (name?: string) => void)(id);
    }
  }, [GlobalRefs, id]);

  // Prefill and rule-based disable/hide
  React.useEffect((): void => {
    const ensureInOptions = (vals: string[]): string[] => clampToExisting(vals, options);

    if (!isLockedRef.current) {
      if (!didInitRef.current) {
        if (FormMode !== 3) {
          const initArr = Array.isArray(starterValue)
            ? (starterValue as Array<string | number>).map(toKey)
            : [toKey(starterValue)];
          setSelectedOptions(ensureInOptions(initArr));
        }
      } else {
        let raw: unknown;
        if (isLookup) {
          if (multiSelect) {
            // read array of objects -> take LookupId
            const mLookup = (FormData as Record<string, unknown> | undefined)?.[id];
            if (Array.isArray(mLookup)) {
              raw = (mLookup as Array<Record<string, unknown>>).map(v =>
                (v as { LookupId?: unknown }).LookupId
              );
            } else {
              raw = [];
            }
          } else {
            raw = (FormData as Record<string, unknown> | undefined)?.[`${id}LookupId`];
          }
        } else {
          raw = (FormData as Record<string, unknown> | undefined)?.[id];
        }

        const arr = ensureInOptions(normalizeToStringArray(raw));
        setSelectedOptions(arr);
      }
      didInitRef.current = true;
    } else {
      const clamped = ensureInOptions(selectedOptions);
      if (clamped.length !== selectedOptions.length) {
        setSelectedOptions(clamped);
      }
    }

    if (FormMode === 4) {
      setIsDisabled(true);
      isLockedRef.current = true;
      const labels = selectedOptions.map(k => keyToText.get(k) ?? k);
      setDisplayOverride(labels.join('; '));
    } else {
      // names and types aligned to your FormFieldsProps
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
        // cast through unknown because UserInfo is not indexable
        curUserList:   ((curUserInfo ?? {}) as unknown as Record<string, unknown>),
        curField:      displayName,
        formStateData: formStateKeys,    // string[]
        listColumns:   listColumnsArray, // string[]
      };

      const results =
        (formFieldsSetup(formFieldProps) as Array<{ isDisabled?: boolean; isHidden?: boolean }>) ?? [];

      if (results.length > 0) {
        for (let i = 0; i < results.length; i += 1) {
          const d = results[i]?.isDisabled;
          if (d !== undefined) {
            setIsDisabled(!!d);
            setDefaultIsDisable(!!d);
          }
          const h = results[i]?.isHidden;
          if (h !== undefined) setIsHidden(!!h);
        }
      }

      if (!isLockedRef.current && isDisabled) {
        const labels = selectedOptions.map(k => keyToText.get(k) ?? k);
        setDisplayOverride(labels.join('; '));
      }
    }

    reportError('');
    setTouched(false);
  }, [
    FormData,
    FormMode,
    id,
    displayName,
    options,
    isLookup,
    multiSelect,
    AllDisableFields,
    AllHiddenFields,
    userBasedPerms,
    curUserInfo,
    listCols,
    selectedOptions,
    isDisabled,
    keyToText,
    starterValue,
    reportError,
  ]);

  const validate = React.useCallback((): string => {
    if (isRequired && selectedOptions.length === 0) return REQUIRED_MSG;
    return '';
  }, [isRequired, selectedOptions]);

  // Commit: send null when empty; numbers for lookup
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

  // Proper v9 types — fixes your onOptionSelect type error
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

  // Semicolon-joined labels for display
  const selectedLabels = selectedOptions.map(k => keyToText.get(k) ?? k);
  const joinedText = selectedLabels.join('; ');
  const visibleText = displayOverride || joinedText;
  const triggerText = visibleText || '';
  const triggerPlaceholder = triggerText || (placeholder || '');
  const hasError = !!error;

  // Build class and attributes so parent CSS gray-out continues to work
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
            multiselect={multiSelect}
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
            ref={elemRef}   // HTMLButtonElement ref (✅ matches v9)
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



