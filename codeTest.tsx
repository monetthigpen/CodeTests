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

/** Option type */
type Opt = { key: string | number; text: string };

/** Props */
interface DropdownProps {
  id: string;
  displayName: string;
  options: Opt[];
  starterValue?: string | number | Array<string | number>;
  isRequired?: boolean;
  placeholder?: string;
  /** v8 prop name */
  multiSelect?: boolean;
  /** v9 prop name */
  multiselect?: boolean;
  /** set to "lookup" if this is a lookup field */
  fieldType?: string; // 'lookup'
  className?: string;
  description?: string;
  disabled?: boolean;
  submitting?: boolean;

  /**
   * Which API are you committing to?
   *  - 'graph' => <InternalName>LookupId with number | number[]
   *  - 'rest'  => <InternalName>Id with number | {results:number[]}
   *
   * Defaults to 'graph'.
   */
  apiFlavor?: 'graph' | 'rest';
}

const REQUIRED_MSG = 'This is a required field and cannot be blank!';

const toKey = (k: unknown): string => (k == null ? '' : String(k));

/** Normalizes unknown into string[] (semicolon-delimited strings supported) */
function normalizeToStringArray(input: unknown): string[] {
  if (input == null) return [];

  // SharePoint classic multi format: { results: any[] }
  if (typeof input === 'object' && Array.isArray((input as any).results)) {
    return ((input as any).results as unknown[]).map(toKey);
  }

  if (Array.isArray(input)) {
    // Could be an array of primitives or objects
    if (input.length && typeof input[0] === 'object') {
      return (input as any[]).map(v => {
        const o = v as any;
        return toKey(o?.LookupId ?? o?.Id ?? o);
      });
    }
    return (input as (string | number)[]).map(toKey);
  }

  if (typeof input === 'string' && input.includes(';')) {
    return input
      .split(';')
      .map(s => toKey(s.trim()))
      .filter(Boolean);
  }

  return [toKey(input)];
}

/** For multi-lookup when reading from FormData in various shapes */
function extractMultiLookupRaw(v: unknown): string[] {
  // REST style {results:number[]}
  if (v && typeof v === 'object' && Array.isArray((v as any).results)) {
    return (v as any).results.map(toKey);
  }

  // Graph style number[]
  if (Array.isArray(v)) {
    if (v.length && typeof v[0] === 'object' && v[0] !== null) {
      return (v as any[]).map(x => toKey((x as any).LookupId ?? (x as any).Id ?? x));
    }
    return (v as (number | string)[]).map(toKey);
  }

  return normalizeToStringArray(v);
}

/** Keeps only values present in options */
function clampToExisting(values: string[], opts: Opt[]): string[] {
  const allowed = new Set(opts.map(o => toKey(o.key)));
  return values.filter(v => allowed.has(v));
}

export default function DropdownComponent(props: DropdownProps): JSX.Element {
  const {
    id,
    starterValue,
    displayName,
    isRequired: requiredProp = false,
    placeholder,
    multiSelect = false,
    multiselect, // v9 mirror (we’ll still drive behavior from multiSelect)
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

  // Context
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

  // State
  const [isRequired, setIsRequired] = React.useState<boolean>(!!requiredProp);
  const [isDisabled, setIsDisabled] = React.useState<boolean>(!!disabledProp);
  const [isHidden, setIsHidden] = React.useState<boolean>(false);
  const [selectedOptions, setSelectedOptions] = React.useState<string[]>([]);
  const [error, setError] = React.useState<string>('');
  const [touched, setTouched] = React.useState<boolean>(false);
  const [displayOverride, setDisplayOverride] = React.useState<string>('');

  const isLockedRef = React.useRef<boolean>(false);
  const didInitRef = React.useRef<boolean>(false);

  // Fast map for key->text
  const keyToText = React.useMemo(() => {
    const m = new Map<string, string>();
    for (const o of options) m.set(toKey(o.key), o.text);
    return m;
  }, [options]);

  /** Decide the commit field name once (depends on API flavor and lookup-ness) */
  const targetFieldName = React.useMemo(() => {
    if (isLookup) {
      return apiFlavor === 'graph' ? `${id}LookupId` : `${id}Id`;
    }
    return id;
  }, [apiFlavor, id, isLookup]);

  /** Send error to global handler using the actual commit field name */
  const reportError = React.useCallback(
    (msg: string) => {
      setError(msg || '');
      GlobalErrorHandle?.(targetFieldName, msg || null);
    },
    [GlobalErrorHandle, targetFieldName]
  );

  // Mirror prop changes
  React.useEffect(() => {
    setIsRequired(!!requiredProp);
    setIsDisabled(!!disabledProp);
  }, [requiredProp, disabledProp]);

  // Initialize / refresh selection from FormData + starterValue
  React.useEffect(() => {
    if (isLockedRef.current) return;

    // First pass: prefer explicit starterValue (e.g., when creating)
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
          (FormData as any)?.[id] ??
          (FormData as any)?.[`${id}LookupId`] ??
          (FormData as any)?.[`${id}Id`];
        raw = extractMultiLookupRaw(mv);
      } else {
        const sv =
          (FormData as any)?.[`${id}LookupId`] ?? // Graph
          (FormData as any)?.[`${id}Id`] ??       // REST
          (FormData as any)?.[id];
        raw = sv;
      }
    } else {
      raw =
        isMulti
          ? (FormData as any)?.[id] ??
            (FormData as any)?.[`${id}Id`] ??
            (FormData as any)?.[`${id}LookupId`]
          : (FormData as any)?.[id];
    }

    const normalized = clampToExisting(normalizeToStringArray(raw), options);
    setSelectedOptions(normalized);
  }, [FormData, FormMode, id, isLookup, isMulti, options, starterValue]);

  // Lock & show display text when submitting or in display mode
  React.useEffect(() => {
    if (submitting || FormMode === 4) {
      setIsDisabled(true);
      isLockedRef.current = true;
      const labels = selectedOptions.map(k => keyToText.get(k) ?? k);
      setDisplayOverride(labels.join('; '));
    }
  }, [FormMode, submitting, selectedOptions, keyToText]);

  // Field-level disable/hide rules
  React.useEffect(() => {
    if (FormMode === 4) return;

    const formFieldProps: FormFieldsProps = {
      disableList: AllDisableFields,
      HiddenList: AllHiddenFields,
      UserBasedList: userBasedPerms,
      curUserList: curUserInfo,
      curField: displayName,
      formStateData: FormData,
      ListColumns: listCols,
    } as any;

    const results = formFieldsSetup(formFieldProps) || [];
    if (results.length > 0) {
      for (let i = 0; i < results.length; i++) {
        if (results[i].isDisabled !== undefined) setIsDisabled(results[i].isDisabled);
        if (results[i].isHidden !== undefined) setIsHidden(results[i].isHidden);
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
  const commitValue = React.useCallback(() => {
    const err = validate();
    reportError(err);

    if (isLookup) {
      const nums = selectedOptions
        .map(k => Number(k))
        .filter(n => Number.isFinite(n));

      if (apiFlavor === 'graph') {
        // GRAPH: <InternalName>LookupId
        (GlobalFormData as any)[targetFieldName] = isMulti
          ? nums                                // number[]
          : (nums[0] ?? null);                  // number | null
      } else {
        // REST: <InternalName>Id
        (GlobalFormData as any)[targetFieldName] = isMulti
          ? { results: nums }                   // { results: number[] }
          : (nums[0] ?? null);                  // number | null
      }
    } else {
      // Non-lookup
      if (isMulti) {
        (GlobalFormData as any)[targetFieldName] =
          apiFlavor === 'graph'
            ? selectedOptions                   // string[]
            : { results: selectedOptions };     // REST: {results:string[]}
      } else {
        (GlobalFormData as any)[targetFieldName] = selectedOptions[0] ?? null;
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

  // Handle UI events (exact Fluent v9 types; selectedOptions/optionValue are optional)
  const handleOptionSelect = React.useCallback(
    (_e: SelectionEvents, data: OptionOnSelectData) => {
      const next = (data.selectedOptions ?? []).map(toKey);
      setSelectedOptions(next);
      if (!touched) reportError(isRequired && next.length === 0 ? REQUIRED_MSG : '');
    },
    [isRequired, reportError, touched]
  );

  const handleBlur = () => {
    setTouched(true);
    commitValue();
  };

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
