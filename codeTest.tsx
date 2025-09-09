import * as React from 'react';
import { Dropdown, IDropdownOption } from '@fluentui/react';
import { Field } from '@fluentui/react-components';
import { DynamicFormContext } from './DynamicFormContext';

export interface DropdownFieldProps {
  id: string;
  displayName: string;
  starterValue?: string | number | Array<string | number>;
  isRequired?: boolean;
  disabled?: boolean;
  placeholder?: string;
  multiSelect?: boolean;   // main prop
  multiselect?: boolean;   // alias
  options: IDropdownOption[];
  fieldType?: string;
  className?: string;
  description?: string;
  submitting?: boolean;    // <— added
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
    multiSelect,
    multiselect,
    options,
    //fieldType,
    className,
    description,
  } = props;

  const isMulti = !!(multiselect ?? multiSelect);

  // NOTE: removed isSubmitting from here
  const { FormData, GlobalFormData, FormMode, GlobalErrorHandle } =
    React.useContext(DynamicFormContext);

  const [isRequired, setIsRequired] = React.useState<boolean>(!!requiredProp);
  const [isDisabled, setIsDisabled] = React.useState<boolean>(!!disabledProp);

  React.useEffect(() => {
    setIsRequired(!!requiredProp);
    setIsDisabled(!!disabledProp);
  }, [requiredProp, disabledProp]);

  // Keep internal selection as strings to satisfy Dropdown types
  const [selectedKeys, setSelectedKeys] = React.useState<string[]>([]);      // for multi
  const [selectedKey, setSelectedKey] = React.useState<string | null>(null); // for single

  const [error, setError] = React.useState<string>('');
  const [touched, setTouched] = React.useState<boolean>(false);

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

  // Prefill: New (8) vs Edit/View
  React.useEffect(() => {
    if (FormMode == 8) {
      if (isMulti) {
        const initArr = toKeyArray(starterValue);
        setSelectedKeys(initArr);
        setSelectedKey(null);
        // GlobalFormData removed here
      } else {
        const init = starterValue != null ? toKey(starterValue as any) : ''; // eslint-disable-line @typescript-eslint/no-explicit-any
        setSelectedKey(init || null);
        setSelectedKeys([]);
        // GlobalFormData removed here
      }
    } else {
      const existing = (FormData ? (props.fieldType=="lookup"? (FormData as any)[`${id}Id`] : (FormData as any)[id]) : undefined); // eslint-disable-line @typescript-eslint/no-explicit-any
      if (isMulti) {
        const arr = toKeyArray(existing);
        setSelectedKeys(arr);
        setSelectedKey(null);
        // GlobalFormData removed here
      } else {
        const k = existing != null ? toKey(existing) : '';
        setSelectedKey(k || null);
        setSelectedKeys([]);
        // GlobalFormData removed here
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

  const handleChange = (
    _e: React.FormEvent<HTMLElement | HTMLDivElement>,
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

  return (
    <Field
      label={displayName}
      required={isRequired}
      validationMessage={hasError ? error : undefined}
      validationState={hasError ? 'error' : undefined}
      submitting={!!props.submitting}            // <— use prop, not context
    >
      <Dropdown
        id={id}
        placeholder={placeholder}
        multiSelect={isMulti}
        disabled={isDisabled}
        inlinePopup={true}                        // <— set true
        // IMPORTANT: use the correct prop for each mode,
        selectedKeys={isMulti ? selectedKeys : undefined}
        selectedKey={!isMulti ? (selectedKey ?? undefined) : undefined}
        options={options}
        onChange={handleChange}
        onBlur={handleBlur}
        className={className}
      />
      {description !== '' && <div className="descriptionText">{description}</div>}
    </Field>
  );
}
