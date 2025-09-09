import * as React from 'react';
import { Field, Dropdown, Option } from '@fluentui/react-components';
import { DynamicFormContext } from './DynamicFormContext';

// Minimal option shape compatible with your data
type OptionItem = { key: string | number; text: string };

export interface DropdownFieldProps {
  id: string;
  displayName: string;
  starterValue?: string | number | Array<string | number>;
  isRequired?: boolean;
  disabled?: boolean;
  placeholder?: string;
  multiSelect?: boolean;   // v8-style name
  multiselect?: boolean;   // v9 prop name
  options: OptionItem[];
  fieldType?: string;      // e.g., "lookup"
  className?: string;
  description?: string;
  submitting?: boolean;
}

const REQUIRED_MSG = 'This is a required field and cannot be blank!';

// ---------- helpers ---------------------------------------------------------

// Accept anything and coerce to string safely
const toKey = (k: unknown): string => (k == null ? '' : String(k));

/** Accepts many backend shapes and returns array<string> of option keys */
function normalizeToStringArray(input: unknown): string[] {
  if (input == null) return [];

  // { results: [...] }
  if (Array.isArray((input as any)?.results)) {
    return ((input as any).results as unknown[]).map(toKey);
  }

  // Array of primitives or objects
  if (Array.isArray(input)) {
    const arr = input as unknown[];
    // If objects like [{ Id: 3 }, { Id: 5 }]
    if (arr.length > 0 && typeof arr[0] === 'object' && arr[0] !== null) {
      return arr.map((o: any) => toKey(o?.Id ?? o?.id ?? o?.Key ?? o?.value ?? o)); // eslint-disable-line @typescript-eslint/no-explicit-any
    }
    return arr.map(toKey);
  }

  // Semicolon-delimited "1;2;3"
  if (typeof input === 'string' && input.includes(';')) {
    return input.split(';').map(s => toKey(s.trim())).filter(Boolean);
  }

  // Single object like { Id: 3, Title: 'x' }
  if (typeof input === 'object') {
    const o: any = input; // eslint-disable-line @typescript-eslint/no-explicit-any
    return [toKey(o?.Id ?? o?.id ?? o?.Key ?? o?.value ?? o)];
  }

  // Single primitive
  return [toKey(input)];
}

const toKeyArray = (v: unknown): string[] =>
  v == null ? [] : Array.isArray(v) ? v.map(toKey) : [toKey(v)];

// Ensure the selected values actually exist in options (v9 shows only values with a matching Option)
function clampToExisting(values: string[], opts: OptionItem[]): string[] {
  if (!values.length) return values;
  const allowed = new Set(opts.map(o => toKey(o.key)));
  return values.filter(v => allowed.has(v));
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
  } = props;

  const isMulti = !!(multiselect ?? multiSelect);

  const { FormData, GlobalFormData, FormMode, GlobalErrorHandle } =
    React.useContext(DynamicFormContext);

  const isSubmitting = !!props.submitting;

  const [isRequired, setIsRequired] = React.useState<boolean>(!!requiredProp);
  const [isDisabled, setIsDisabled] = React.useState<boolean>(!!disabledProp);

  const [selectedKeys, setSelectedKeys] = React.useState<string[]>([]);      // multi
  const [selectedKey, setSelectedKey] = React.useState<string | null>(null); // single

  const [error, setError] = React.useState<string>('');
  const [touched, setTouched] = React.useState<boolean>(false);

  React.useEffect(() => {
    setIsRequired(!!requiredProp);
    setIsDisabled(!!disabledProp);
  }, [requiredProp, disabledProp]);

  // ---------- Prefill (Edit/View + wait for data/options) -------------------
  React.useEffect(() => {
    // Disable while submitting
    if (props.submitting === true) {
      setIsDisabled(true);
    }

    const ensureInOptions = (vals: string[]) => clampToExisting(vals, options);

    if (FormMode == 8) {
      // New form: use starterValue
      if (isMulti) {
        const initArr = ensureInOptions(toKeyArray(starterValue));
        setSelectedKeys(initArr);
        setSelectedKey(null);
      } else {
        const init = starterValue != null ? toKey(starterValue) : '';
        const clamped = ensureInOptions(init ? [init] : []);
        setSelectedKey(clamped[0] ?? null);
        setSelectedKeys([]);
      }
    } else {
      // Edit/View: derive from FormData, handling lookups + various shapes
      const raw =
        FormData
          ? (fieldType === 'lookup'
              ? (FormData as any)[`${id}Id`] // eslint-disable-line @typescript-eslint/no-explicit-any
              : (FormData as any)[id])       // eslint-disable-line @typescript-eslint/no-explicit-any
          : undefined;

      if (isMulti) {
        const arr = clampToExisting(normalizeToStringArray(raw), options);
        setSelectedKeys(arr);
        setSelectedKey(null);
      } else {
        const arr = clampToExisting(normalizeToStringArray(raw), options);
        const first = arr.length ? arr[0] : '';
        setSelectedKey(first || null);
        setSelectedKeys([]);
      }
    }

    setError('');
    setTouched(false);
    GlobalErrorHandle(id, null);

    // IMPORTANT: include options and FormData so defaults populate when either arrives
  }, [
    FormData,
    FormMode,
    starterValue,
    options,
    props.submitting,
    fieldType,
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
    GlobalFormData(id, isMulti ? selectedKeys : selectedKey);
    GlobalErrorHandle(id, err);
  }, [validate, GlobalFormData, GlobalErrorHandle, id, isMulti, selectedKeys, selectedKey]);

  // v9 selection handler
  const handleOptionSelect = (
    _e: unknown,
    data: { optionValue?: string | number; selectedOptions: (string | number)[] }
  ) => {
    if (isMulti) {
      const next = (data.selectedOptions ?? []).map(toKey);
      setSelectedKeys(next);
      if (touched) setError(isRequired && next.length === 0 ? REQUIRED_MSG : '');
    } else {
      const next = data.optionValue != null ? toKey(data.optionValue) : '';
      setSelectedKey(next || null);
      if (touched) setError(isRequired && !next ? REQUIRED_MSG : '');
    }
  };

  const handleBlur = () => {
    setTouched(true);
    commitValue();
  };

  const hasError = !!error;

  // Allow custom 'submitting' prop on Field (optional)
  const FieldAny = Field as any;

  const selectedOptions = isMulti
    ? selectedKeys
    : selectedKey
      ? [selectedKey]
      : [];

  return (
    <FieldAny
      label={displayName}
      required={isRequired}
      validationMessage={hasError ? error : undefined}
      validationState={hasError ? 'error' : undefined}
      submitting={isSubmitting}
    >
      <Dropdown
        id={id}
        placeholder={placeholder}
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

      {description !== '' && (
        <div className="descriptionText">{description}</div>
      )}
    </FieldAny>
  );
}


