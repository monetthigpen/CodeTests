import * as React from 'react';
import { Field, Dropdown, Option } from '@fluentui/react-components';
import { DynamicFormContext } from './DynamicFormContext';
import formFieldsSetup, { FormFieldsProps } from './formFieldBased';

interface DropdownProps {
  id: string;
  starterValue?: string | number | Array<string | number>;
  displayName: string;
  isRequired: boolean;
  placeholder?: string;
  multiSelect?: boolean;  // v8 prop
  multiselect?: boolean;  // v9 prop
  fieldType?: string;     // "lookup" to commit under `${id}Id` as numbers
  options: { key: string | number; text: string }[];
  className?: string;
  description?: string;
  disabled?: boolean;
  submitting?: boolean;
}

const REQUIRED_MSG = 'This is a required field and cannot be blank!';

const toKey = (k: unknown): string => (k == null ? '' : String(k));

export default function DropdownComponent(props: DropdownProps): JSX.Element {
  const {
    id,
    starterValue,
    displayName,
    isRequired,
    placeholder,
    multiSelect,
    multiselect,
    fieldType,
    options,
    className,
    description,
    disabled,
    submitting,
  } = props;

  const [localVal, setLocalVal] = React.useState<string | string[]>('');
  const [error, setError] = React.useState<string>('');
  const [isDisabled, setIsDisabled] = React.useState<boolean>(disabled ?? false);
  const [isHidden, setIsHidden] = React.useState<boolean>(false);

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

  // Handle blur
  const handleBlur = (e: React.FocusEvent<HTMLElement>): void => {
    if ((Array.isArray(localVal) && localVal.length === 0) || localVal === '') {
      if (isRequired) {
        setError(REQUIRED_MSG);
        GlobalErrorHandle(id, REQUIRED_MSG);
      } else {
        setError('');
        // SharePoint requires null for empty values
        // eslint-disable-next-line @rushstack/no-new-null
        GlobalErrorHandle(id, null);
      }
    } else {
      setError('');
      // eslint-disable-next-line @rushstack/no-new-null
      GlobalErrorHandle(id, null);
    }
  };

  // Handle change
  const handleChange = (e: React.ChangeEvent<HTMLSelectElement>): void => {
    const value = e.target.value;
    setLocalVal(value);

    if (value === '') {
      if (isRequired) {
        setError(REQUIRED_MSG);
        GlobalErrorHandle(id, REQUIRED_MSG);
      } else {
        setError('');
        // eslint-disable-next-line @rushstack/no-new-null
        GlobalErrorHandle(id, null);
      }
    } else {
      setError('');
      GlobalFormData(id, value);
      // eslint-disable-next-line @rushstack/no-new-null
      GlobalErrorHandle(id, null);
    }
  };

  // Initial value load
  React.useEffect((): void => {
    if (FormMode === 4 || FormMode === 6) {
      const fldInternalName = id;
      if (FormData !== undefined) {
        const fieldValue = FormData[fldInternalName];
        if (fieldValue !== null && fieldValue !== undefined) {
          setLocalVal(fieldValue);
        }
      }
    } else if (FormMode === 8 && starterValue !== undefined) {
      setLocalVal(starterValue as string);
      GlobalFormData(id, starterValue);
    }
  }, [FormMode, FormData, id, starterValue, GlobalFormData]);

  // Disable/hide logic
  React.useEffect((): void => {
    if (FormMode === 4) {
      setIsDisabled(true);
    } else {
      const formFieldProps: FormFieldsProps = {
        disabledList: AllDisableFields,
        hiddenList: AllHiddenFields,
        userBasedList: userBasedPerms,
        curUserList: curUserInfo,
        curField: displayName,
        formStateData: FormData,
        listColumns: listCols,
      };

      const results = formFieldsSetup(formFieldProps);
      if (results.length > 0) {
        for (let i = 0; i < results.length; i++) {
          if (results[i].isDisabled !== undefined) {
            setIsDisabled(results[i].isDisabled);
          }
          if (results[i].isHidden !== undefined) {
            setIsHidden(results[i].isHidden);
          }
        }
      }
    }
  }, [FormMode, AllDisableFields, AllHiddenFields, userBasedPerms, curUserInfo, displayName, FormData, listCols]);

  // Submit disable
  React.useEffect((): void => {
    if (submitting === true) {
      setIsDisabled(true);
    }
  }, [submitting]);

  return (
    <div style={{ display: isHidden ? 'none' : 'block' }}>
      <Field
        label={displayName}
        {...(isRequired && { required: true })}
        validationMessage={error || undefined}
        validationState={error ? 'error' : undefined}
      >
        <Dropdown
          id={id}
          className={className}
          placeholder={placeholder}
          value={localVal}
          disabled={isDisabled}
          inlinePopup={true}
          onBlur={handleBlur}
          onChange={handleChange}
        >
          {options.map((o) => (
            <Option key={toKey(o.key)} value={toKey(o.key)}>
              {o.text}
            </Option>
          ))}
        </Dropdown>
        {description && (
          <div className="descriptionText">{description}</div>
        )}
      </Field>
    </div>
  );
}






