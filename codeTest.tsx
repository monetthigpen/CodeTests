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

const toKey = (k: unknown): string => (k == null ? '' : String(k));

function normalizeToStringArray(input: unknown): string[] {
  if (input == null) return [];
  if (Array.isArray((input as any)?.results)) {
    return ((input as any).results as unknown[]).map(toKey);
  }
  if (Array.isArray(input)) {
    const arr = input as unknown[];
    if (arr.length && typeof arr[0] === 'object' && arr[0] !== null) {
      return arr.map((o: any) => toKey(o?.Id ?? o?.id ?? o?.Key ?? o?.value ?? o)); // eslint-disable-line
    }
    return arr.map(toKey);
  }
  if (typeof input === 'string' && input.includes(';')) {
    return input.split(';').map(s => toKey(s.trim())).filter(Boolean);
  }
  if (typeof input === 'object') {
    const o: any = input; // eslint-disable-line
    return [toKey(o?.Id ?? o?.id ?? o?.Key ?? o?.value ?? o)];
  }
  return [toKey(input)];
}

function clampToExisting(values: string[], opts: OptionItem[]): string[] {
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
    submitting,
  } = props;

  const isMulti = !!(multiselect ?? multiSelect);

  const { FormData, GlobalFormData, FormMode, GlobalErrorHandle } =
    React.useContext(DynamicFormContext);

  const [isRequired, setIsRequired] = React.useState<boolean>(!!requiredProp);
  const [isDisabled, setIsDisabled] = React.useState<boolean>(!!disabledProp);
  const [selectedKeys, setSelectedKeys] = React.useState<string[]>([]);      // multi
  const [selectedKey, setSelectedKey] = React.useState<string | null>(null); // single
  const [error, setError] = React.useState<string>('');
  const [touched, setTouched] = React.useState<boolean>(false);

  React.useEffect(() => {
    setIsRequired(!!requiredProp);
    setIsDisabled(submitting ? true : !!disabledProp);
  }, [requiredProp, disabledProp, submitting]);

  // ---------- Prefill (Edit/View + wait for data/options) -------------------
  React.useEffect(() => {
    if (FormMode == 8) {
      // New form: use starterValue
      if (isMulti) {
        const initArr = clampToExisting(
          starterValue != null
            ? (Array.isArray(starterValue) ? starterValue.map(toKey) : [toKey(starterValue)])
            : [],
          options
        );
        setSelectedKeys(initArr);
        setSelectedKey(null);
      } else {
        const init = starterValue != null ? toKey(starterValue) : '';
        const clamped = clampToExisting(init ? [init] : [], options);
        setSelectedKey(clamped[0] ?? null); // <-- never ""
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
        const arr = clampToExisting(normalizeToStringArray(raw), options);
        setSelectedKeys(arr);
        setSelectedKey(null);
      } else {
        const arr = clampToExisting(normalizeToStringArray(raw), options);
        setSelectedKey(arr[0] ?? null); // <-- never ""
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

    // Single-select: use null instead of "" when empty
    const valueForCommit =
      isMulti
        ? selectedKeys // [] when empty
        : (selectedKey ? selectedKey : null); // <-- null not ""

    GlobalFormData(id, valueForCommit);
    GlobalErrorHandle(id, err);
  }, [validate, GlobalFormData, GlobalErrorHandle, id, isMulti, selectedKeys, selectedKey]);

  // v9 selection handler (use null when deselected)
  const handleOptionSelect = (
    _e: unknown,
    data: { optionValue?: string | number; selectedOptions: (string | number)[] }
  ) => {
    if (isMulti) {
      const next = (data.selectedOptions ?? []).map(toKey);
      setSelectedKeys(next);
      if (touched) setError(isRequired && next.length === 0 ? REQUIRED_MSG : '');
    } else {
      // When nothing is selected, v9 may pass undefined. Convert to null.
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
  const keyToText = React.useMemo(() => {
    const map = new Map<string, string>();
    for (const o of options) map.set(toKey(o.key), o.text);
    return map;
  }, [options]);

  const selectedOptions = isMulti
    ? selectedKeys
    : selectedKey
      ? [selectedKey]
      : [];

  const displayText =
    isMulti
      ? (selectedKeys.length
          ? selectedKeys.map(k => keyToText.get(k) ?? k).join(', ')
          : '')
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


