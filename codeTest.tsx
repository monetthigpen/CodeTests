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

type Opt = { key: string | number; text: string };

interface DropdownProps {
  id: string;
  displayName: string;
  options: Opt[];

  starterValue?: string | number | Array<string | number>;
  isRequired?: boolean;
  placeholder?: string;

  /** v8 style */
  multiSelect?: boolean;
  /** v9 style */
  multiselect?: boolean;

  /** mark as lookup, otherwise inferred by id suffixes if you prefer */
  fieldType?: 'lookup' | string;

  className?: string;
  description?: string;
  disabled?: boolean;
  submitting?: boolean;

  /** 'graph' => LookupId field, 'rest' => Id / {results:[]} (you can swap when you wire REST) */
  apiFlavor?: 'graph' | 'rest';
}

/* -------------------- helpers (typed & lint-clean) -------------------- */

const REQUIRED_MSG = 'This is a required field and cannot be blank!';

const toKey = (k: unknown): string => (k == null ? '' : String(k));

/** Remove trailing Id/LookupId so we can append the right suffix cleanly */
const baseName = (name: string): string => name.replace(/(Lookup)?Id$/i, '');

const clampToExisting = (values: string[], opts: Opt[]): string[] => {
  const allowed = new Set(opts.map(o => toKey(o.key)));
  return values.filter(v => allowed.has(v));
};

const normalizeToStringArray = (input: unknown): string[] => {
  if (input == null) return [];

  // REST multi: {results: []}
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
};

const extractMultiLookupRaw = (v: unknown): string[] => {
  if (
    v &&
    typeof v === 'object' &&
    Array.isArray((v as { results?: unknown[] }).results)
  ) {
    return ((v as { results: unknown[] }).results).map(toKey);
  }
  if (Array.isArray(v)) {
    const arr = v as unknown[];
    if (arr.length > 0 && typeof arr[0] === 'object' && arr[0] !== null) {
      return (arr as Array<Record<string, unknown>>).map(x =>
        toKey((x as { LookupId?: unknown; Id?: unknown }).LookupId ?? (x as { Id?: unknown }).Id ?? x)
      );
    }
    return (arr as Array<number | string>).map(toKey);
  }
  return normalizeToStringArray(v);
};

/* -------------------- component -------------------- */

export default function DropdownComponent(props: DropdownProps): JSX.Element {
  const {
    id,
    starterValue,
    displayName,
    isRequired: requiredProp = false,
    placeholder,
    multiSelect = false,
    multiselect,
    fieldType,
    options,
    className,
    description,
    disabled: disabledProp = false,
    submitting = false,
    apiFlavor = 'graph',
  } = props;

  const isLookup = fieldType === 'lookup';
  const isMulti = !!multiSelect || !!multiselect;

  // Use your context as-is; its default values make properties defined
  const {
    FormData,
    GlobalFormData,
    FormMode,
    GlobalErrorHandle,
    AllDisableFields,
    AllHiddenFields,
    userBasedPerms,
    curUserInfo,
    listCols,
  } = React.useContext(DynamicFormContext);

  // state
  const [isRequired, setIsRequired] = React.useState<boolean>(!!requiredProp);
  const [isDisabled, setIsDisabled] = React.useState<boolean>(!!disabledProp);
  const [isHidden, setIsHidden] = React.useState<boolean>(false);
  const [selectedOptions, setSelectedOptions] = React.useState<string[]>([]);
  const [error, setError] = React.useState<string>('');
  const [touched, setTouched] = React.useState<boolean>(false);
  const [displayOverride, setDisplayOverride] = React.useState<string>('');

  const isLockedRef = React.useRef<boolean>(false);
  const didInitRef = React.useRef<boolean>(false);

  // key -> text
  const keyToText = React.useMemo<Map<string, string>>(() => {
    const m = new Map<string, string>();
    for (const o of options) m.set(toKey(o.key), o.text);
    return m;
  }, [options]);

  // where we actually write in the outbound payload
  const targetFieldName = React.useMemo<string>(() => {
    if (isLookup) {
      const base = baseName(id);
      return apiFlavor === 'graph' ? `${base}LookupId` : `${base}Id`;
    }
    return id;
  }, [apiFlavor, id, isLookup]);

  const reportError = React.useCallback((msg: string): void => {
    setError(msg || '');
    GlobalErrorHandle?.(targetFieldName, msg || null);
  }, [GlobalErrorHandle, targetFieldName]);

  React.useEffect((): void => {
    setIsRequired(!!requiredProp);
    setIsDisabled(!!disabledProp);
  }, [requiredProp, disabledProp]);

  // initialize + reflect FormData
  React.useEffect((): void => {
    if (isLockedRef.current) return;

    const base = baseName(id);

    if (!didInitRef.current) {
      if (FormMode !== 3) {
        const initArr = Array.isArray(starterValue)
          ? (starterValue as Array<string | number>).map(toKey)
          : [toKey(starterValue)];
        setSelectedOptions(clampToExisting(initArr, options));
      }
      didInitRef.current = true;
      return;
    }

    let raw: unknown;

    if (isLookup) {
      if (isMulti) {
        const mv =
          (FormData as Record<string, unknown> | undefined)?.[base] ??
          (FormData as Record<string, unknown> | undefined)?.[`${base}LookupId`] ??
          (FormData as Record<string, unknown> | undefined)?.[`${base}Id`] ??
          (FormData as Record<string, unknown> | undefined)?.[id];
        raw = extractMultiLookupRaw(mv);
      } else {
        const sv =
          (FormData as Record<string, unknown> | undefined)?.[`${base}LookupId`] ??
          (FormData as Record<string, unknown> | undefined)?.[`${base}Id`] ??
          (FormData as Record<string, unknown> | undefined)?.[base] ??
          (FormData as Record<string, unknown> | undefined)?.[id];
        raw = sv;
      }
    } else {
      raw =
        (FormData as Record<string, unknown> | undefined)?.[base] ??
        (FormData as Record<string, unknown> | undefined)?.[id];
    }

    const normalized = clampToExisting(normalizeToStringArray(raw), options);
    setSelectedOptions(normalized);
  }, [FormData, FormMode, id, isLookup, isMulti, options, starterValue]);

  // lock & show text if submitting or read-only
  React.useEffect((): void => {
    if (submitting || FormMode === 4) {
      setIsDisabled(true);
      isLockedRef.current = true;
      const labels = selectedOptions.map(k => keyToText.get(k) ?? k);
      setDisplayOverride(labels.join('; '));
    }
  }, [FormMode, submitting, selectedOptions, keyToText]);

  // field-level disable/hide
  React.useEffect((): void => {
    if (FormMode === 4) return;

    // Your FormFieldsProps in your project expects these keys and (often) array/object types
    const formStateKeys: string[] = Array.isArray(FormData)
      ? (FormData as string[])
      : Object.keys((FormData ?? {}) as Record<string, unknown>);

    const formFieldProps: FormFieldsProps = {
      disabledList:  (AllDisableFields ?? {}) as Record<string, unknown>,
      hiddenList:    (AllHiddenFields ?? {}) as Record<string, unknown>,
      userBasedList: (userBasedPerms ?? {}) as Record<string, unknown>,
      curUserList:   (curUserInfo ?? {}) as Record<string, unknown>,
      curField:      displayName,
      formStateData: formStateKeys, // string[]
      listColumns:   (Array.isArray(listCols) ? listCols : []) as unknown[], // keep as array form
    };

    const results =
      (formFieldsSetup(formFieldProps) as Array<{ isDisabled?: boolean; isHidden?: boolean }>) ?? [];

    if (results.length > 0) {
      for (let i = 0; i < results.length; i += 1) {
        if (typeof results[i].isDisabled === 'boolean') setIsDisabled(results[i].isDisabled);
        if (typeof results[i].isHidden === 'boolean') setIsHidden(results[i].isHidden);
      }
    }

    if (!isLockedRef.current && isDisabled) {
      const labels = selectedOptions.map(k => keyToText.get(k) ?? k);
      setDisplayOverride(labels.join('; '));
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [
    AllDisableFields,
    AllHiddenFields,
    FormData,
    FormMode,
    curUserInfo,
    displayName,
    isDisabled,
    keyToText,
    listCols,
    selectedOptions,
    userBasedPerms,
  ]);

  // validation
  const validate = React.useCallback((): string => {
    return isRequired && selectedOptions.length === 0 ? REQUIRED_MSG : '';
  }, [isRequired, selectedOptions]);

  // commit into GlobalFormData
  const commitValue = React.useCallback((): void => {
    const err = validate();
    reportError(err);

    if (isLookup) {
      const nums = selectedOptions
        .map(k => Number(k))
        .filter((n): n is number => Number.isFinite(n));

      if (apiFlavor === 'graph') {
        (GlobalFormData as (n: string, v: unknown) => void)(
          targetFieldName,
          isMulti ? nums : (nums[0] ?? null)
        );
      } else {
        (GlobalFormData as (n: string, v: unknown) => void)(
          targetFieldName,
          isMulti ? { results: nums } : (nums[0] ?? null)
        );
      }
    } else {
      if (apiFlavor === 'graph') {
        (GlobalFormData as (n: string, v: unknown) => void)(
          targetFieldName,
          isMulti ? selectedOptions : (selectedOptions[0] ?? null)
        );
      } else {
        (GlobalFormData as (n: string, v: unknown) => void)(
          targetFieldName,
          isMulti ? { results: selectedOptions } : (selectedOptions[0] ?? null)
        );
      }
    }

    const labels = selectedOptions.map(k => keyToText.get(k) ?? k);
    setDisplayOverride(labels.join('; '));
  }, [
    apiFlavor,
    GlobalFormData,
    isLookup,
    isMulti,
    keyToText,
    reportError,
    selectedOptions,
    targetFieldName,
    validate,
  ]);

  // handlers (typed)
  const handleOptionSelect = React.useCallback(
    (_e: SelectionEvents, data: OptionOnSelectData): void => {
      const next = (data.selectedOptions ?? []).map(toKey);
      setSelectedOptions(next);
      if (!touched) {
        reportError(isRequired && next.length === 0 ? REQUIRED_MSG : '');
      }
    },
    [isRequired, reportError, touched]
  );

  const handleBlur = React.useCallback((): void => {
    setTouched(true);
    commitValue();
  }, [commitValue]);

  // render helpers
  const selectedLabels = selectedOptions.map(k => keyToText.get(k) ?? k);
  const joinedText = selectedLabels.join('; ');
  const triggerText = displayOverride || joinedText;
  const triggerPlaceholder = triggerText || placeholder || '';

  const disabledClass = isDisabled ? 'is-disabled' : '';
  const rootClassName = [className, disabledClass].filter(Boolean).join(' ');

  return (
    <div style={{ display: isHidden ? 'none' : 'block' }}>
      <Field
        label={displayName}
        required={isRequired}
        validationMessage={error || undefined}
        validationState={error ? 'error' : undefined}
      >
        {isDisabled ? (
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
          >
            {options.map((o: Opt) => (
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


