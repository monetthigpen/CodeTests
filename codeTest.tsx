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

/** Option type for the dropdown */
type Opt = { key: string | number; text: string };

interface DropdownProps {
  /** Base SharePoint internal name (no suffix). If you pass ...Id/LookupId, we’ll strip it. */
  id: string;
  displayName: string;
  options: Opt[];

  starterValue?: string | number | Array<string | number>;
  isRequired?: boolean;
  placeholder?: string;

  /** v8 prop name for multi */
  multiSelect?: boolean;
  /** v9 prop name for multi */
  multiselect?: boolean;

  /** Set to "lookup" if this is a lookup field (auto-detected if id ends w/ Id/LookupId) */
  fieldType?: string;

  className?: string;
  description?: string;
  disabled?: boolean;
  submitting?: boolean;

  /**
   * API payload shape:
   *  - 'graph' (default): <InternalName>LookupId with number | number[]
   *  - 'rest'           : <InternalName>Id with number | {results:number[]}
   */
  apiFlavor?: 'graph' | 'rest';
}

type GlobalFormDataShape = Record<string, unknown>;

const REQUIRED_MSG = 'This is a required field and cannot be blank!';
const toKey = (k: unknown): string => (k == null ? '' : String(k));

/** Strip trailing Id/LookupId so we can re-append the correct suffix cleanly */
function baseName(name: string): string {
  return name.replace(/(Lookup)?Id$/i, '');
}

/** Keep only keys that exist in options */
function clampToExisting(values: string[], opts: Opt[]): string[] {
  const allowed = new Set(opts.map(o => toKey(o.key)));
  return values.filter(v => allowed.has(v));
}

/** Normalize any unknown value into string[] (handles {results:[]}, arrays, semicolons) */
function normalizeToStringArray(input: unknown): string[] {
  if (input == null) return [];

  // REST multi: { results: any[] }
  if (typeof input === 'object' && input !== null && Array.isArray((input as { results?: unknown[] }).results)) {
    return ((input as { results: unknown[] }).results).map(toKey);
  }

  if (Array.isArray(input)) {
    if (input.length > 0 && typeof input[0] === 'object' && input[0] !== null) {
      // Array of objects (edit/display)
      return (input as Array<Record<string, unknown>>).map(o =>
        toKey((o as { LookupId?: unknown; Id?: unknown }).LookupId ?? (o as { Id?: unknown }).Id ?? o)
      );
    }
    return (input as Array<string | number>).map(toKey);
  }

  if (typeof input === 'string' && input.includes(';')) {
    return input.split(';').map(s => toKey(s.trim())).filter(Boolean);
  }

  return [toKey(input)];
}

/** For multi-lookup when reading from FormData in various shapes */
function extractMultiLookupRaw(v: unknown): string[] {
  // REST: {results:number[]}
  if (v && typeof v === 'object' && Array.isArray((v as { results?: unknown[] }).results)) {
    return ((v as { results: unknown[] }).results).map(toKey);
  }
  // Graph: number[] or array of objects
  if (Array.isArray(v)) {
    if (v.length > 0 && typeof v[0] === 'object' && v[0] !== null) {
      return (v as Array<Record<string, unknown>>).map(x =>
        toKey((x as { LookupId?: unknown; Id?: unknown }).LookupId ?? (x as { Id?: unknown }).Id ?? x)
      );
    }
    return (v as Array<number | string>).map(toKey);
  }
  return normalizeToStringArray(v);
}

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

  // Treat as lookup if explicitly set OR id ends with Id/LookupId
  const isLookup: boolean = fieldType === 'lookup' || /(Lookup)?Id$/i.test(id);
  const isMulti: boolean = !!multiSelect || !!multiselect;

  // Use your actual DynamicFormContext shape (from your screenshot)
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

  const [isRequired, setIsRequired] = React.useState<boolean>(!!requiredProp);
  const [isDisabled, setIsDisabled] = React.useState<boolean>(!!disabledProp);
  const [isHidden, setIsHidden] = React.useState<boolean>(false);
  const [selectedOptions, setSelectedOptions] = React.useState<string[]>([]);
  const [error, setError] = React.useState<string>('');
  const [touched, setTouched] = React.useState<boolean>(false);
  const [displayOverride, setDisplayOverride] = React.useState<string>('');

  const isLockedRef = React.useRef<boolean>(false);
  const didInitRef = React.useRef<boolean>(false);

  // key -> text map
  const keyToText = React.useMemo<Map<string, string>>(() => {
    const m = new Map<string, string>();
    for (const o of options) m.set(toKey(o.key), o.text);
    return m;
  }, [options]);

  /** Actual field name we will commit to (Graph vs REST) */
  const targetFieldName = React.useMemo<string>(() => {
    if (isLookup) {
      const base = baseName(id);
      return apiFlavor === 'graph' ? `${base}LookupId` : `${base}Id`;
    }
    return id;
  }, [apiFlavor, id, isLookup]);

  /** Report error against the real commit field name */
  const reportError = React.useCallback((msg: string): void => {
    setError(msg || '');
    GlobalErrorHandle?.(targetFieldName, msg || null);
  }, [GlobalErrorHandle, targetFieldName]);

  // React to prop changes
  React.useEffect(() => {
    setIsRequired(!!requiredProp);
    setIsDisabled(!!disabledProp);
  }, [requiredProp, disabledProp]);

  // Initialize / refresh from FormData + starterValue
  React.useEffect(() => {
    if (isLockedRef.current) return;

    const base = baseName(id);

    // First pass: use starterValue (e.g., create form)
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

    // Subsequent passes: reflect FormData
    let raw: unknown;

    if (isLookup) {
      if (isMulti) {
        const mv =
          (FormData as any)?.[base] ??
          (FormData as any)?.[`${base}LookupId`] ??
          (FormData as any)?.[`${base}Id`] ??
          (FormData as any)?.[id];
        raw = extractMultiLookupRaw(mv);
      } else {
        const sv =
          (FormData as any)?.[`${base}LookupId`] ?? // Graph
          (FormData as any)?.[`${base}Id`] ??       // REST
          (FormData as any)?.[base] ??              // plain
          (FormData as any)?.[id];
        raw = sv;
      }
    } else {
      raw = isMulti
        ? (FormData as any)?.[base] ?? (FormData as any)?.[id]
        : (FormData as any)?.[base] ?? (FormData as any)?.[id];
    }

    const normalized = clampToExisting(normalizeToStringArray(raw), options);
    setSelectedOptions(normalized);
  }, [FormData, FormMode, id, isLookup, isMulti, options, starterValue]);

  // Lock & show display text when submitting or in read-only mode
  React.useEffect(() => {
    if (submitting || FormMode === 4) {
      setIsDisabled(true);
      isLockedRef.current = true;
      const labels = selectedOptions.map(k => keyToText.get(k) ?? k);
      setDisplayOverride(labels.join('; '));
    }
  }, [FormMode, submitting, selectedOptions, keyToText]);

  // ---- Field-level disable/hide rules — keys match FormFieldsProps ----
  React.useEffect(() => {
    if (FormMode === 4) return;

    const formFieldProps: FormFieldsProps = {
      disabledList:  (AllDisableFields ?? {}) as Record<string, any>,
      hiddenList:    (AllHiddenFields ?? {}) as Record<string, any>,
      userBasedList: (userBasedPerms ?? {}) as Record<string, any>,
      curUserList:   (curUserInfo ?? {}) as Record<string, any>,
      curField:      displayName,
      formStateData: (FormData ?? {}) as Record<string, any>,
      listColumns:   (listCols ?? {}) as Record<string, any>,
    };

    const results =
      (formFieldsSetup(formFieldProps) as Array<{ isDisabled?: boolean; isHidden?: boolean }>) ?? [];

    if (results.length > 0) {
      for (let i = 0; i < results.length; i++) {
        if (results[i].isDisabled !== undefined) setIsDisabled(results[i].isDisabled as boolean);
        if (results[i].isHidden   !== undefined) setIsHidden(results[i].isHidden as boolean);
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

  // Validation
  const validate = React.useCallback((): string => {
    return isRequired && selectedOptions.length === 0 ? REQUIRED_MSG : '';
  }, [isRequired, selectedOptions]);

  // Commit into GlobalFormData in the exact shape expected by the chosen API
  const commitValue = React.useCallback((): void => {
    const err = validate();
    reportError(err);

    const dest = GlobalFormData as GlobalFormDataShape;

    if (isLookup) {
      const nums = selectedOptions
        .map(k => Number(k))
        .filter((n): n is number => Number.isFinite(n));

      if (apiFlavor === 'graph') {
        // Graph: <InternalName>LookupId
        dest[targetFieldName] = isMulti ? nums : (nums[0] ?? null);
      } else {
        // REST: <InternalName>Id
        dest[targetFieldName] = isMulti ? { results: nums } : (nums[0] ?? null);
      }
    } else {
      // Non-lookup (choice/text/etc.)
      if (isMulti) {
        dest[targetFieldName] = apiFlavor === 'graph'
          ? selectedOptions
          : { results: selectedOptions }; // REST shape
      } else {
        dest[targetFieldName] = selectedOptions[0] ?? null;
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

  // Handle UI events (Fluent v9 types; properties may be undefined)
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

  // Render helpers
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
          // Keep gray-out visuals but show the selected text
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

