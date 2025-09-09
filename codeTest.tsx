import * as React from 'react';
import { Field, Dropdown, Option } from '@fluentui/react-components';
import { DynamicFormContext } from './DynamicFormContext';

// Minimal option shape compatible with your existing data
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

// Normalize any key to a string
const toKey = (k: string | number | null | undefined): string =>
  k == null ? '' : String(k);

const toKeyArray = (v: unknown): string[] =>
  v == null ? [] : Array.isArray(v) ? v.map(toKey) : [toKey(v as string | number)];

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
    className,
    description,
  } = props;

  // v9 prop is "multiselect"; keep support for either spelling
  const isMulti = !!(multiselect ?? multiSelect);

  const { FormData, GlobalFormData, FormMode, GlobalErrorHandle } =
    React.useContext(DynamicFormContext);

  // derive submitting from prop (not from context)
  const isSubmitting = !!props.submitting;

  const [isRequired, setIsRequired] = React.useState<boolean>(!!requiredProp);
  const [isDisabled, setIsDisabled] = React.useState<boolean>(!!disabledProp);

  // Keep internal selection as strings for consistency
  const [selectedKeys, setSelectedKeys] = React.useState<string[]>([]);      // multi
  const [selectedKey, setSelectedKey] = React.useState<string | null>(null); // single

  const [error, setError] = React.useState<string>('');
  const [touched, setTouched] = React.useState<boolean>(false);

  React.useEffect(() => {
    setIsRequired(!!requiredProp);
    setIsDisabled(!!disabledProp);
  }, [requiredProp, disabledProp]);

  // Prefill: New (8) vs Edit/View — NO GlobalFormData writes here
  React.useEffect(() => {
    if (FormMode == 8) {
      if (isMulti) {
        const initArr = toKeyArray(starterValue);
        setSelectedKeys(initArr);
        setSelectedKey(null);
      } else {
        const init = starterValue != null ? toKey(starterValue as any) : ''; // eslint-disable-line @typescript-eslint/no-explicit-any
        setSelectedKey(init || null);
        setSelectedKeys([]);
      }
    } else {
      const existing = (FormData
        ? (props.fieldType === 'lookup'
            ? (FormData as any)[`${id}Id`]
            : (FormData as any)[id])
        : undefined) as unknown; // eslint-disable-line @typescript-eslint/no-explicit-any

      if (isMulti) {
        const arr = toKeyArray(existing);
        setSelectedKeys(arr);
        setSelectedKey(null);
      } else {
        const k = existing != null ? toKey(existing as any) : ''; // eslint-disable-line @typescript-eslint/no-explicit-any
        setSelectedKey(k || null);
        setSelectedKeys([]);
      }
    }

    if (props.submitting === true) {
      setIsDisabled(true);
    }

    setError('');
    setTouched(false);
    GlobalErrorHandle(id, null);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [props.submitting]);

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
    // Commit to your external form state on blur/change as before
    GlobalFormData(id, isMulti ? selectedKeys : selectedKey);
    GlobalErrorHandle(id, err);
  }, [validate, GlobalFormData, GlobalErrorHandle, id, isMulti, selectedKeys, selectedKey]);

  // v9 selection handler
  const handleOptionSelect = (
    _e: unknown,
    data: { optionValue?: string | number; selectedOptions: (string | number)[] }
  ) => {
    // v9 provides selectedOptions already; we still normalize to strings
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

  // Cast to allow custom 'submitting' prop on Field per your requirement
  const FieldAny = Field as any;

  // Build selectedOptions for v9
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


