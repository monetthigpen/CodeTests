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
  fieldType?: string;
  className?: string;
  description?: string;
  submitting?: boolean;
}

const REQUIRED_MSG = 'This is a required field and cannot be blank!';

// ---- helpers ---------------------------------------------------------------

// Accept anything and coerce to string safely
const toKey = (k: unknown): string =>
  k == null ? '' : String(k);

/** Accepts many backend shapes and returns array<string> of option keys */
function normalizeToStringArray(input: unknown): string[] {
  if (input == null) return [];

  // { results: [...] }
  if (Array.isArray((input as any).results)) {
    return ((input as any).results as unknown[]).map(toKey);
  }

  // Already an array
  if (Array.isArray(input)) {
    return (input as unknown[]).map(toKey);
  }

  // Semicolon-delimited "1;2;3"
  if (typeof input === 'string' && input.includes(';')) {
    return input.split(';').map(s => toKey(s.trim())).filter(Boolean);
  }

  // Single value
  return [toKey(input)];
}

const toKeyArray = (v: unknown): string[] =>
  v == null ? [] : Array.isArray(v) ? v.map(toKey) : [toKey(v)];

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

  // ---- Prefill (Edit/View support) ----------------------------------------
  React.useEffect(() => {
    if (FormMode == 8) {
      // New form: use starterValue
      if (isMulti) {
        const initArr = toKeyArray(starterValue);
        setSelectedKeys(initArr);
        setSelectedKey(null);
      } else {
        const init = starterValue != null ? toKey(starterValue) : '';
        setSelectedKey(init || null);
        setSelectedKeys([]);
      }
    } else {
      // Edit/View: derive from FormData
      const raw =
        FormData
          ? (fieldType === 'lookup'
              ? (FormData as any)[`${id}Id`] // eslint-disable-line @typescript-eslint/no-explicit-any
              : (FormData as any)[id])       // eslint-disable-line @typescript-eslint/no-explicit-any
          : undefined;

      if (isMulti) {
        const arr = normalizeToStringArray(raw);
        setSelectedKeys(arr);
        setSelectedKey(null);
      } else {
        const arr = normalizeToStringArray(raw);
        const first = arr.length ? arr[0] : '';
        setSelectedKey(first || null);
        setSelectedKeys([]);
      }
    }

    if (props.submitting === true) {
      setIsDisabled(true);
    }

    setError('');
    setTouched(false);
    GlobalErrorHandle(id, null);
    // include dependencies so it updates when data loads/changes
  }, [FormData, FormMode, starterValue, props.submitting, fieldType, id, isMulti, GlobalErrorHandle]);

  // ---- Validation / Commit -------------------------------------------------
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

  // allow custom 'submitting' prop on Field
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

