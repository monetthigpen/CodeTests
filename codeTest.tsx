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
  multiSelect?: boolean;
  multiselect?: boolean;
  options: OptionItem[];
  fieldType?: string;
  className?: string;
  description?: string;
  submitting?: boolean; // new boolean prop
}

const REQUIRED_MSG = 'This is a required field and cannot be blank!';

const toKey = (k: unknown): string => (k == null ? '' : String(k));

function normalizeToStringArray(input: unknown): string[] {
  if (input == null) return [];
  if (Array.isArray((input as any)?.results)) return ((input as any).results as unknown[]).map(toKey);
  if (Array.isArray(input)) {
    const arr = input as unknown[];
    if (arr.length && typeof arr[0] === 'object' && arr[0] !== null) {
      return arr.map((o: any) => toKey(o?.Id ?? o?.id ?? o?.Key ?? o?.value ?? o));
    }
    return arr.map(toKey);
  }
  if (typeof input === 'string' && input.includes(';')) {
    return input.split(';').map(s => toKey(s.trim())).filter(Boolean);
  }
  if (typeof input === 'object') {
    const o: any = input;
    return [toKey(o?.Id ?? o?.id ?? o?.Key ?? o?.value ?? o)];
  }
  return [toKey(input)];
}

function clampToExisting(values: string[], opts: OptionItem[]): string[] {
  const allowed = new Set(opts.map(o => toKey(o.key)));
  return values.filter(v => allowed.has(v));
}

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
  const { FormData, GlobalFormData, FormMode, GlobalErrorHandle } = React.useContext(DynamicFormContext);

  const [isRequired, setIsRequired] = React.useState<boolean>(!!requiredProp);
  const [isDisabled, setIsDisabled] = React.useState<boolean>(!!disabledProp);
  const [selectedKeys, setSelectedKeys] = React.useState<string[]>([]);
  const [selectedKey, setSelectedKey] = React.useState<string | null>(null);
  const [error, setError] = React.useState<string>('');
  const [touched, setTouched] = React.useState<boolean>(false);

  // react to disabled/required props and submitting flag
  React.useEffect(() => {
    setIsRequired(!!requiredProp);

    if (submitting) {
      // force disable while submitting is true
      setIsDisabled(true);
    } else {
      setIsDisabled(!!disabledProp);
    }
  }, [requiredProp, disabledProp, submitting]);

  // Prefill values
  React.useEffect(() => {
    if (FormMode == 8) {
      if (isMulti) {
        const initArr = clampToExisting(
          starterValue != null ? (Array.isArray(starterValue) ? starterValue.map(toKey) : [toKey(starterValue)]) : [],
          options
        );
        setSelectedKeys(initArr);
        setSelectedKey(null);
      } else {
        const init = starterValue != null ? toKey(starterValue) : '';
        setSelectedKey(init || null);
        setSelectedKeys([]);
      }
    } else {
      const raw = FormData
        ? (fieldType === 'lookup'
            ? (FormData as any)[`${id}Id`]
            : (FormData as any)[id])
        : undefined;

      if (isMulti) {
        const arr = clampToExisting(normalizeToStringArray(raw), options);
        setSelectedKeys(arr);
        setSelectedKey(null);
      } else {
        const arr = clampToExisting(normalizeToStringArray(raw), options);
        setSelectedKey(arr[0] ?? null);
        setSelectedKeys([]);
      }
    }

    setError('');
    setTouched(false);
    GlobalErrorHandle(id, null);
  }, [FormData, FormMode, starterValue, options, fieldType, id, isMulti, GlobalErrorHandle]);

  // validation + commit
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

  const handleOptionSelect = (_: unknown, data: { optionValue?: string | number; selectedOptions: (string | number)[] }) => {
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

  const keyToText = React.useMemo(() => {
    const map = new Map<string, string>();
    for (const o of options) map.set(toKey(o.key), o.text);
    return map;
  }, [options]);

  const selectedOptions = isMulti ? selectedKeys : selectedKey ? [selectedKey] : [];
  const displayText = isMulti
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


