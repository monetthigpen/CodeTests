import * as React from 'react';
import {
  Field,
  Dropdown,
  Option,
  Input,
  type SelectionEvents,
  type OptionOnSelectData,
} from '@fluentui/react-components';
import { DynamicFormContext } from './DynamicFormContext';
import formFieldsSetup, { type FormFieldsProps } from './formFieldBased';

interface DropdownProps {
  id: string;
  starterValue?: string | number | Array<string | number>;
  displayName: string;
  isRequired?: boolean;
  placeholder?: string;
  multiSelect?: boolean;   // v8 prop
  multiselect?: boolean;   // v9 prop
  fieldType?: string;      // 'lookup' => commit under `${id}Id` as numbers
  options: { key: string | number; text: string }[];
  className?: string;
  description?: string;
  disabled?: boolean;
  submitting?: boolean;
}

const REQUIRED_MSG = 'This is a required field and cannot be blank!';

// Keep original nullish semantics
// eslint-disable-next-line eqeqeq
const toKey = (k: unknown): string => (k == null ? '' : String(k));

function normalizeToStringArray(input: unknown): string[] {
  if (input === null || input === undefined) return [];

  // SharePoint REST multi: { results: [] }
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
  // Fluent v9 Dropdown trigger is a button
  const elemRef = React.useRef<HTMLButtonElement | null>(null);

  const {
    id,
    starterValue,
    displayName,
    isRequired: requiredProp = false,
    placeholder,
    multiSelect = false,
    multiselect: multiselectProp = false, // accept both spellings
    fieldType,
    options,
    className,
    description,
    disabled: disabledProp = false,
    submitting = false,
  } = props;

  const isMulti = !!multiSelect || !!multiselectProp;
  const isLookup = fieldType === 'lookup';

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

  const [isRequired, setIsRequired] = React.useState<boolean>(!!requiredProp);
  const [isDisabled, setIsDisabled] = React.useState<boolean>(!!disabledProp);
  const [isHidden, setIsHidden] = React.useState<boolean>(false);
  const [defaultIsDisable, setDefaultIsDisable] = React.useState<boolean>(false);

  // Controlled selection
  const [selectedOptions, setSelectedOptions] = React.useState<string[]>([]);
  const [error, setError] = React.useState<string>('');
  const [touched, setTouched] = React.useState<boolean>(false);

  // Lock/Cache for display text when disabled
  const [displayOverride, setDisplayOverride] = React.useState<string>('');
  const isLockedRef = React.useRef<boolean>(false);
  const didInitRef = React.useRef<boolean>(false);

  const keyToText = React.useMemo((): Map<string, string> => {
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

  // Register ref with context (supports function or map-like forms)
  React.useEffect((): void => {
    if (typeof GlobalRefs === 'function') {
      (GlobalRefs as (elmIntrnName?: string) => void)(id);
    } else {
      (GlobalRefs as unknown as Record<string, unknown>)[id] = elemRef.current ?? undefined;
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
          const clamped = ensureInOptions(initArr);
          if (clamped.length !== selectedOptions.length || clamped.some((v, i) => v !== selectedOptions[i])) {
            setSelectedOptions(clamped);
          }
        } else {
          let raw: unknown;
          if (isLookup) {
            if (isMulti) {
              const mLookup = (FormData as Record<string, unknown> | undefined)?.[id];
              raw = Array.isArray(mLookup)
                ? (mLookup as Array<Record<string, unknown>>).map(v => (v as { LookupId?: unknown }).LookupId)
                : [];
            } else {
              raw = (FormData as Record<string, unknown> | undefined)?.[`${id}LookupId`];
            }
          } else {
            raw = (FormData as Record<string, unknown> | undefined)?.[id];
          }
          const arr = ensureInOptions(normalizeToStringArray(raw));
          if (arr.length !== selectedOptions.length || arr.some((v, i) => v !== selectedOptions[i])) {
            setSelectedOptions(arr);
          }
        }
        didInitRef.current = true;
      } else {
        const clamped = ensureInOptions(selectedOptions);
        if (clamped.length !== selectedOptions.length || clamped.some((v, i) => v !== selectedOptions[i])) {
          setSelectedOptions(clamped);
        }
      }
    }

    if (FormMode === 4) {
      setIsDisabled(true);
      isLockedRef.current = true;
      const labels = selectedOptions.map(k => keyToText.get(k) ?? k);
      setDisplayOverride(labels.join('; '));
    } else {
      const formFieldProps: FormFieldsProps = {
        disabledList:  ((AllDisableFields ?? {}) as unknown as Record<string, unknown>),
        hiddenList:    ((AllHiddenFields ?? {}) as unknown as Record<string, unknown>),
        userBasedList: ((userBasedPerms ?? {}) as unknown as Record<string, unknown>),
        curUserList:   ((curUserInfo ?? {}) as unknown as Record<string, unknown>),
        curField:      displayName,
        // Many implementations expect string[]; pass known keys if object
        formStateData: Array.isArray(FormData)
          ? (FormData as string[])
          : Object.keys((FormData ?? {}) as Record<string, unknown>),
        listColumns:   Array.isArray(listCols) ? (listCols as string[]).map(String) : [],
      } as unknown as FormFieldsProps;

      const results =
        (formFieldsSetup(formFieldProps) as Array<{ isDisabled?: boolean; isHidden?: boolean }>) ?? [];

      if (results.length > 0) {
        for (let i = 0; i < results.length; i++) {
          if (results[i].isDisabled !== undefined) {
            setIsDisabled(Boolean(results[i].isDisabled));
            setDefaultIsDisable(Boolean(results[i].isDisabled));
          }
          if (results[i].isHidden !== undefined) setIsHidden(Boolean(results[i].isHidden));
        }
      }

      if (!isLockedRef.current && isDisabled) {
        const labels = selectedOptions.map(k => keyToText.get(k) ?? k);
        setDisplayOverride(labels.join('; '));
      }
    }

    reportError('');
    setTouched(false);
    // keep broad deps as in your original
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [
    FormData,
    FormMode,
    id,
    displayName,
    options,
    isLookup,
    isMulti,
    AllDisableFields,
    AllHiddenFields,
    userBasedPerms,
    curUserInfo,
    listCols,
    selectedOptions,
    isDisabled,
    starterValue,
    keyToText,
    reportError,
  ]);

  const validate = React.useCallback((): string => {
    if (isRequired && selectedOptions.length === 0) return REQUIRED_MSG;
    return '';
  }, [isRequired, selectedOptions]);

  // Commit: send undefined when empty; numbers for lookup (Graph prefers numbers)
  const commitValue = React.useCallback((): void => {
    const err = validate();
    reportError(err);

    const targetId = isLookup ? `${id}Id` : id;

    if (isLookup) {
      const nums = selectedOptions.map(k => Number(k)).filter(n => Number.isFinite(n));
      if (isMulti) {
        (GlobalFormData as (name: string, value: unknown) => void)(targetId, nums.length === 0 ? undefined : nums);
      } else {
        (GlobalFormData as (name: string, value: unknown) => void)(targetId, nums.length === 0 ? undefined : nums[0]);
      }
    } else {
      (GlobalFormData as (name: string, value: unknown) => void)(
        targetId,
        selectedOptions.length === 0 ? undefined : (isMulti ? selectedOptions : selectedOptions[0])
      );
    }

    const labels = selectedOptions.map(k => keyToText.get(k) ?? k);
    setDisplayOverride(labels.join('; '));
  }, [validate, reportError, GlobalFormData, id, isLookup, isMulti, selectedOptions, keyToText]);

  // Fluent v9 types; no 'any'
  const handleOptionSelect = (_e: SelectionEvents, data: OptionOnSelectData): void => {
    const next = (data.selectedOptions ?? []).map(toKey);
    setSelectedOptions(next);
    if (!touched) reportError(isRequired && next.length === 0 ? REQUIRED_MSG : '');
  };

  const handleBlur = (): void => {
    setTouched(true);
    commitValue();
  };

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
            multiselect={isMulti}
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



