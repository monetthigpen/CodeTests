/* eslint-disable @typescript-eslint/no-explicit-any */
/* eslint-disable @typescript-eslint/explicit-function-return-type */
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

// keep original semantics: treat null OR undefined as empty string
// eslint-disable-next-line eqeqeq
const toKey = (k: unknown): string => (k == null ? '' : String(k));

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

// tiny guards so the big effect won't thrash even with many deps
const arraysEqual = (a: string[], b: string[]): boolean =>
  a.length === b.length && a.every((v, i) => v === b[i]);

export default function DropdownComponent(props: DropdownProps): JSX.Element {
  // NOTE: Fluent v9 Dropdown trigger is a BUTTON → ref should be HTMLButtonElement
  const elemRef = React.useRef<HTMLButtonElement | null>(null); // used to get the DOM element especially to get look up values

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

  const keyToText = React.useMemo(() => {
    const m = new Map<string, string>();
    for (const o of options) m.set(toKey(o.key), o.text);
    return m;
  }, [options]);

  const reportError = React.useCallback((msg: string): void => {
    const targetId = isLookup ? `${id}LookupId` : id;
    setError(msg || '');
    GlobalErrorHandle?.(targetId, msg || null);
  }, [GlobalErrorHandle, id, isLookup]);

  React.useEffect(() => {
    setIsRequired(!!requiredProp);
    setIsDisabled(!!disabledProp);
  }, [requiredProp, disabledProp]);

  // Submitting disables and locks display text
  React.useEffect(() => {
    if (defaultIsDisable === false) {
      setIsDisabled(!!submitting);
    } else {
      setIsDisabled(true);
      isLockedRef.current = true;
      const labels = selectedOptions.map(k => keyToText.get(k) ?? k);
      setDisplayOverride(labels.join('; '));
    }
  }, [submitting, defaultIsDisable, selectedOptions, keyToText]);

  // Register ref (GlobalRefs is a function in your context — your screenshot)
  React.useEffect(() => {
    if (typeof GlobalRefs === 'function') {
      (GlobalRefs as (elmIntrnName?: string) => void)(id);
    } else {
      // fallback if elsewhere it's a map
      (GlobalRefs as unknown as Record<string, unknown>)[id] = elemRef.current ?? undefined;
    }
  }, [GlobalRefs, id]);

  // Prefill and rule-based disable/hide  (structure preserved)
  React.useEffect(() => {
    const ensureInOptions = (vals: string[]) => clampToExisting(vals, options);

    if (!isLockedRef.current) {
      if (!didInitRef.current) {
        if (FormMode !== 3) {
          const initArr =
            Array.isArray(starterValue)
              ? (starterValue as Array<string | number>).map(toKey)
              : [toKey(starterValue)];
          const clamped = ensureInOptions(initArr);
          if (!arraysEqual(selectedOptions, clamped)) setSelectedOptions(clamped);
        }
      } else {
        let raw: any;
        if (isLookup) {
          if (multiSelect) {
            const mLookup = (FormData as any)?.[id];
            if (Array.isArray(mLookup)) {
              raw = (mLookup as Array<Record<string, unknown>>).map(v => (v as { LookupId?: unknown }).LookupId);
            } else {
              raw = [];
            }
          } else {
            raw = (FormData as any)?.[`${id}LookupId`];
          }
        } else {
          raw = (FormData as any)?.[id];
        }
        const arr = ensureInOptions(normalizeToStringArray(raw));
        if (!arraysEqual(selectedOptions, arr)) setSelectedOptions(arr);
      }
      didInitRef.current = true;
    } else {
      const clamped = ensureInOptions(selectedOptions);
      if (!arraysEqual(selectedOptions, clamped)) setSelectedOptions(clamped);
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
        curUserList:   ((curUserInfo ?? {}) as unknown as Record<string, unknown>), // cast via unknown to satisfy indexable
        curField:      displayName,
        formStateData: Array.isArray(FormData)
          ? (FormData as string[])
          : Object.keys((FormData ?? {}) as Record<string, unknown>), // your interface wants string[]
        listColumns:   Array.isArray(listCols) ? (listCols as string[]).map(String) : [], // ensure string[]
      } as unknown as FormFieldsProps;

      const results = (formFieldsSetup(formFieldProps) || []) as Array<{ isDisabled?: boolean; isHidden?: boolean }>;
      if (results.length > 0) {
        for (let i = 0; i < results.length; i++) {
          const d = results[i].isDisabled;
          if (d !== undefined && d !== isDisabled) {
            setIsDisabled(!!d);
            setDefaultIsDisable(!!d);
          }
          const h = results[i].isHidden;
          if (h !== undefined && h !== isHidden) setIsHidden(!!h);
        }
      }

      if (!isLockedRef.current && isDisabled) {
        const labels = selectedOptions.map(k => keyToText.get(k) ?? k);
        setDisplayOverride(labels.join('; '));
      }
    }

    reportError('');
    setTouched(false);
    // keep your original "lots of deps" style, but stop the linter complaining
    // eslint-disable-next-line react-hooks/exhaustive-deps
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
    selectedOptions,   // kept on purpose (original behavior); guarded with arraysEqual
    isDisabled,        // kept on purpose (original behavior); guarded with !== checks
    starterValue,
    keyToText,
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
      const nums = selectedOptions.map(k => Number(k)).filter(n => Number.isFinite(n));
      if (multiSelect) {
        (GlobalFormData as (name: string, value: unknown) => void)(targetId, nums.length === 0 ? null : nums);
      } else {
        (GlobalFormData as (name: string, value: unknown) => void)(targetId, nums.length === 0 ? null : nums[0]);
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

  // Proper Fluent v9 signature (minimal change; fixes your ts error)
  const handleOptionSelect = React.useCallback(
    (_e: SelectionEvents, data: OptionOnSelectData): void => {
      const next = (data.selectedOptions ?? []).map(toKey);
      if (!arraysEqual(selectedOptions, next)) setSelectedOptions(next);
      if (!touched) reportError(isRequired && next.length === 0 ? REQUIRED_MSG : '');
    },
    [isRequired, reportError, touched, selectedOptions]
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


