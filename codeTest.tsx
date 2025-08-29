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

  const [selectedKeys, setSelectedKeys] = React.useState<(string | number)[]>([]);
  const [error, setError] = React.useState<string>('');
  const [touched, setTouched] = React.useState<boolean>(false);

  const validate = React.useCallback(
    (vals: (string | number)[]): string => {
      if (isRequired && vals.length === 0) return REQUIRED_MSG;
      return '';
    },
    [isRequired]
  );

  const commitValue = React.useCallback(
    (vals: (string | number)[]) => {
      const err = validate(vals);
      setError(err);
      GlobalErrorHandle(id, err || null);
      GlobalFormData(id, isMulti ? vals : vals[0] ?? '');
    },
    [id, isMulti, validate, GlobalErrorHandle, GlobalFormData]
  );

  // Prefill starter / FormData
  React.useEffect(() => {
    const normalize = (v: any): (string | number)[] =>
      v == null ? [] : Array.isArray(v) ? v : [v];

    if (FormMode === 8) {
      const initial = normalize(starterValue);
      setSelectedKeys(initial);
      GlobalFormData(id, isMulti ? initial : initial[0] ?? '');
    } else {
      const existing = normalize(FormData ? (FormData as any)[id] : undefined);
      setSelectedKeys(existing);
      GlobalFormData(id, isMulti ? existing : existing[0] ?? '');
    }
    setError('');
    setTouched(false);
    GlobalErrorHandle(id, null);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [FormMode, starterValue, id, isMulti]);

  // Handlers
  const handleChange = (
    _e: React.FormEvent<HTMLDivElement>,
    option?: IDropdownOption
  ) => {
    if (!option) return;

    let next: (string | number)[] = [];

    if (isMulti) {
      if (option.selected) {
        next = [...selectedKeys, option.key];
      } else {
        next = selectedKeys.filter(k => k !== option.key);
      }
    } else {
      next = [option.key];
    }

    setSelectedKeys(next);
    if (touched) setError(validate(next));
  };

  const handleBlur = () => {
    setTouched(true);
    commitValue(selectedKeys);
  };

  const hasError = !!error;

  const dropdownStyles: Partial<IDropdownStyles> = {
    root: { width: '100%' }
  };

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
        selectedKeys={selectedKeys}
        options={options}
        styles={dropdownStyles}
        onChange={handleChange}
        onBlur={handleBlur}
      />
    </Field>
  );
}

