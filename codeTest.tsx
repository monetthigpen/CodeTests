import * as React from 'react';
import {
  Dropdown,
  IDropdownOption,
  IDropdownStyles
} from '@fluentui/react';
import { Field, useId } from '@fluentui/react-components';
import { DynamicFormContext } from './DynamicFormContext';

export interface DropdownFieldProps {
  id: string;
  displayName: string;
  starterValue?: string | number | Array<string | number>;
  isRequired?: boolean;
  disabled?: boolean;
  placeholder?: string;
  multiselect?: boolean;  // main prop
  multiSelect?: boolean;  // alias
  options: IDropdownOption[];
}

const REQUIRED_MSG = 'This is a required field and cannot be blank!';

// Normalize any key to a string (Fluent v8 accepts string keys)
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
    multiselect,
    multiSelect,
    options
  } = props;

  const isMulti = !!(multiselect ?? multiSelect);

  const { FormData, GlobalFormData, FormMode, GlobalErrorHandle } =
    React.useContext(DynamicFormContext);

  const inputId = useId('dropdown');

  const [isRequired, setIsRequired] = React.useState<boolean>(!!requiredProp);
  const [isDisabled, setIsDisabled] = React.useState<boolean>(!!disabledProp);
  React.useEffect(() => {
    setIsRequired(!!requiredProp);
    setIsDisabled(!!disabledProp);
  }, [requiredProp, disabledProp]);

  // Keep internal selection as strings to satisfy Dropdown types
  const [selectedKeys, setSelectedKeys] = React.useState<string[]>([]);   // for multi
  const [selectedKey, setSelectedKey]   = React.useState<string | null>(null); // for single

  const [error, setError] = React.useState<string>('');
  const [touched, setTouched] = React.useState<boolean>(false);

  const validate = React.useCallback((): string => {
    if (!isRequired) return '';
    if (isMulti) return selectedKeys.length === 0 ? REQUIRED_MSG : '';
    return !selectedKey ? REQUIRED_MSG : '';
  }, [isRequired, isMulti, selectedKeys.length, selectedKey]);

  const commitValue = React.useCallback(() => {
    const err = validate();
    setError(err);
    GlobalErrorHandle(id, err || null);
    GlobalFormData(id, isMulti ? selectedKeys : (selectedKey ?? ''));
  }, [validate, GlobalErrorHandle, GlobalFormData, id, isMulti, selectedKeys, selectedKey]);

  // Prefill: New (8) vs Edit/View
  React.useEffect(() => {
    if (FormMode === 8) {
      if (isMulti) {
        const initArr = toKeyArray(starterValue);
        setSelectedKeys(initArr);
        setSelectedKey(null);
        GlobalFormData(id, initArr);
      } else {
        const init = starterValue != null ? toKey(starterValue as any) : '';
        setSelectedKey(init || null);
        setSelectedKeys([]);
        GlobalFormData(id, init || '');
      }
    } else {
      const existing = (FormData ? (FormData as any)[id] : undefined);
      if (isMulti) {
        const arr = toKeyArray(existing);
        setSelectedKeys(arr);
        setSelectedKey(null);
        GlobalFormData(id, arr);
      } else {
        const k = existing != null ? toKey(existing) : '';
        setSelectedKey(k || null);
        setSelectedKeys([]);
        GlobalFormData(id, k || '');
      }
    }
    setError('');
    setTouched(false);
    GlobalErrorHandle(id, null);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [FormMode, starterValue, id, isMulti]);

  const handleChange = (
    _e: React.FormEvent<HTMLDivElement>,
    option?: IDropdownOption
  ) => {
    if (!option) return;
    const k = toKey(option.key);

    if (isMulti) {
      const next = option.selected
        ? [...selectedKeys, k]
        : selectedKeys.filter(x => x !== k);
      setSelectedKeys(next);
      if (touched) setError(isRequired && next.length === 0 ? REQUIRED_MSG : '');
    } else {
      setSelectedKey(k);
      if (touched) setError(isRequired && !k ? REQUIRED_MSG : '');
    }
  };

  const handleBlur = () => {
    setTouched(true);
    commitValue();
  };

  const hasError = !!error;

  const dropdownStyles: Partial<IDropdownStyles> = { root: { width: '100%' } };

  return (
    <Field
      label={displayName}
      required={isRequired}
      validationMessage={hasError ? error : undefined}
      validationState={hasError ? 'error' : undefined}
      size="medium"
    >
      <Dropdown
        id={inputId}
        placeholder={placeholder}
        multiSelect={isMulti}
        disabled={isDisabled}
        // IMPORTANT: use the correct prop for each mode,
        // and pass homogenous types (string[]) to satisfy TS
        selectedKeys={isMulti ? selectedKeys : undefined}
        selectedKey={!isMulti ? (selectedKey ?? undefined) : undefined}
        options={options}
        styles={dropdownStyles}
        onChange={handleChange}
        onBlur={handleBlur}
      />
    </Field>
  );
}
