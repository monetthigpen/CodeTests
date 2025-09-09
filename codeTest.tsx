import * as React from 'react';
import { Field, Dropdown, Option } from '@fluentui/react-components';
import { DynamicFormContext } from './DynamicFormContext';

type OptionItem = { key: string | number; text: string };

export interface DropdownFieldProps {
  id: string;
  displayName: string;
  starterValue?: string | number | Array<string | number>;
  isRequired?: boolean;
  disabled?: boolean;
  placeholder?: string;
  multiSelect?: boolean;     // v8-style name
  multiselect?: boolean;     // v9 prop name
  options: OptionItem[];
  fieldType?: string;        // "lookup" to send numeric Id(s)
  className?: string;
  description?: string;
  submitting?: boolean;
}

const REQUIRED_MSG = 'This is a required field and cannot be blank!';

// ---------- helpers ---------------------------------------------------------

const toKey = (k: unknown): string => (k == null ? '' : String(k));

/** normalize incoming backend shapes into array<string> of option keys */
function normalizeToStringArray(input: unknown): string[] {
  if (input == null) return [];

  // { results: [...] }
  if (Array.isArray((input as any)?.results)) {
    return ((input as any).results as unknown[]).map(toKey);
  }

  // Array of primitives or objects
  if (Array.isArray(input)) {
    const arr = input as unknown[];
    if (arr.length && typeof arr[0] === 'object' && arr[0] !== null) {
      return arr.map((o: any) => toKey(o?.Id ?? o?.id ?? o?.Key ?? o?.value ?? o)); // eslint-disable-line
    }
    return arr.map(toKey);
  }

  // Semicolon-delimited "1;2;3"
  if (typeof input === 'string' && input.includes(';')) {
    return input.split(';').map(s => toKey(s.trim())).filter(Boolean);
  }

  // Single object like { Id: 3, Title: 'x' }
  if (typeof input === 'object') {
    const o: any = input; // eslint-disable-line
    return [toKey(o?.Id ?? o?.id ?? o?.Key ?? o?.value ?? o)];
  }

  // Single primitive
  return [toKey(input)];
}

/** restrict values to those present in options */
function clampToExisting(values: string[], opts: OptionItem[]): string[] {
  const allowed = new Set(opts.map(o => toKey(o.key)));
  return values.filter(v => allowed.has(v));
}

// Build key<->text and key<->number maps once per options change
function useOptionMaps(options: OptionItem[]) {
  return React.useMemo(() => {
    const keyToText = new Map<string, string>();
    const keyToNumber = new Map<string, number>();
    const numberToKey = new Map<number, string>();

    for (const o of options) {
      const keyStr = toKey(o.key);
      keyToText.set(keyStr, o.text);
      // record numeric version if possible (for lookup commits)
      const maybeNum = typeof o.key === 'number' ? o.key : Number.isFinite(Number(keyStr)) ? Number(keyStr) : NaN;
      if (!Number.isNaN(maybeNum)) {
        keyToNumber.set(keyStr, maybeNum);
        numberToKey.set(maybeNum, keyStr);
      }
    }
    return { keyToText, keyToNumber, numberToKey };
  }, [options]);
}

// ---------------------------------------------------------------------------

export default function DropdownField(props: DropdownFieldProps): JSX.Element {
  const {
    id,
    displayName,
    starterValue,
    isRequired: requiredProp,
    disabled: disabledProp,
    placeholder,
    multiSelect,
    multiselect,
    options,
    fieldType,
    className,
    description,
    submitting,
  } = props;

  const isMulti = !!(multiselect ?? multiSelect);
  const isLookup = fieldType === 'lookup';

  const { FormData, GlobalFormData, FormMode, GlobalErrorHandle } =
    React.useContext(DynamicFormContext);

  const [isRequired, setIsRequired] = React.useState<boolean>(!!requiredProp);
  const [isDisabled, setIsDisabled] = React.useState<boolean>(!!disabledProp);
  const [selectedKeys, setSelectedKeys] = React.useState<string[]>([]);      // multi
  const [selectedKey, setSelectedKey] = React.useState<string | null>(null); // single
  const [error, setError] = React.useState<string>('');
  const [touched, setTouched] = React.useState<boolean>(false);

  const { keyToText, keyToNumber, numberToKey } = useOptionMaps(options);

  // react to required/disabled/submitting flags
  React.useEffect(() => {
    setIsRequired(!!requiredProp);
    setIsDisabled(submitting ? true : !!disabledProp);
  }, [requiredProp, disabledProp, submitting]);

  // ---------- Prefill (Edit/View + wait for data/options) -------------------
  React.useEffect(() => {
    // Disable while submitting
    if (submitting === true) setIsDisabled(true);

    const ensureInOptions = (vals: string[]) => clampToExisting(vals, options);

    if (FormMode == 8) {
      // New form: use starterValue (could be numbers or strings)
      if (isMulti) {
        const initArr = ensureInOptions(
          starterValue != null
            ? (Array.isArray(starterValue) ? starterValue.map(toKey) : [toKey(starterValue)])
            : []
        );
        setSelectedKeys(initArr);
        setSelectedKey(null);
      } else {
        const init = starterValue != null ? toKey(starterValue) : '';
        const clamped = ensureInOptions(init ? [init] : []);
        setSelectedKey(clamped[0] ?? null);
        setSelectedKeys([]);
      }
    } else {
      // Edit/View
      const raw = FormData
        ? (isLookup
            ? (FormData as any)[`${id}Id`] // preferred: numeric id(s)
            : (FormData as any)[id])
        : undefined;

      if (isMulti) {
        // raw might be number[], {results}, [{Id}], string[] etc.
        const arr = ensureInOptions(normalizeToStringArray(raw));
        setSelectedKeys(arr);
        setSelectedKey(null);
      } else {
        const arr = ensureInOptions(normalizeToStringArray(raw));
        setSelectedKey(arr[0] ?? null);
        setSelectedKeys([]);
      }
    }

    setError('');
    setTouched(false);
    GlobalErrorHandle(id, null);
  }, [
    FormData,
    FormMode,
    starterValue,
    options,
    submitting,
    isLookup,
    id,
    isMulti,
    GlobalErrorHandle,
  ]);

  // ---------- Validation / Commit ------------------------------------------
  const validate = React.useCallback((): string => {
    if (isRequired) {
      if (isMulti && selectedKeys.length === 0) return REQUIRED_MSG;
      if (!isMulti && !selectedKey) return REQUIRED_MSG;
    }
    return '';
  }, [isRequired, isMulti, selectedKeys, selectedKey]);

  const commitValue = React.useCallback(() => {
    const err = validate();
    setError(err);

    // Build payloads:
    if (isLookup) {
      // Send numeric IDs for lookups
      const valueForCommit = isMulti
        ? selectedKeys.map(k => keyToNumber.get(k)).filter((n): n is number => typeof n === 'number')
        : (selectedKey ? keyToNumber.get(selectedKey) ?? null : null);

      GlobalFormData(id, valueForCommit);
    } else {
      // Non-lookup: keep strings (but null when empty for single)
      const valueForCommit = isMulti ? selectedKeys : (selectedKey ? selectedKey : null);
      GlobalFormData(id, valueForCommit);
    }

    GlobalErrorHandle(id, err);
  }, [validate, GlobalFormData, GlobalErrorHandle, id, isMulti, isLookup, selectedKeys, selectedKey, keyToNumber]);

  // v9 selection handler (state keeps string keys)
  const handleOptionSelect = (_: unknown, data: { optionValue?: string | number; selectedOptions: (string | number)[] }) => {
    if (isMulti) {
      const next = (data.selectedOptions ?? []).map(toKey);
      setSelectedKeys(next);
      if (touched) setError(isRequired && next.length === 0 ? REQUIRED_MSG : '');
    } else {
      const nextVal = data.optionValue != null ? toKey(data.optionValue) : null;
      setSelectedKey(nextVal);
      if (touched) setError(isRequired && !nextVal ? REQUIRED_MSG : '');
    }
  };

  const handleBlur = () => {
    setTouched(true);
    commitValue();
  };

  // ----- Trigger text so the button shows selection(s) ----------------------
  const selectedOptions = isMulti
    ? selectedKeys
    : selectedKey
      ? [selectedKey]
      : [];

  const displayText =
    isMulti
      ? (selectedKeys.length ? selectedKeys.map(k => keyToText.get(k) ?? k).join(', ') : '')
      : (selectedKey ? (keyToText.get(selectedKey) ?? selectedKey) : '');

  const effectivePlaceholder = displayText || placeholder;

  const hasError = !!error;
  const FieldAny = Field as any;

  return (
    <FieldAny
      label={displayName}
      required={isRequired}
      validationMessage={hasError ? error : undefined}
      validationState={hasError ? 'error' : undefined}
      submitting={!!submitting}
    >
      <Dropdown
        id={id}
        placeholder={effectivePlaceholder}
        multiselect={isMulti}
        disabled={isDisabled}
        inlinePopup
        selectedOptions={selectedOptions}
        onOptionSelect={handleOptionSelect}
        onBlur={handleBlur}
        className={className}
      >
        {options.map(o => (
          <Option key={toKey(o.key)} value={toKey(o.key)}>
            {o.text}
          </Option>
        ))}
      </Dropdown>

      {description !== '' && <div className="descriptionText">{description}</div>}
    </FieldAny>
  );
}


