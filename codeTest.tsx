import * as React from 'react';
import { Field, Dropdown, Option, Input, DropdownOnOptionSelectData } from '@fluentui/react-components';
import { DynamicFormContext } from './DynamicFormContext';
import formFieldsSetup, { FormFieldsProps } from './formFieldBased';

interface OptionItem { key: string | number; text: string; }

interface DropdownProps {
  id: string;
  displayName: string;
  options: OptionItem[];
  starterValue?: string | number | Array<string | number>;
  isRequired?: boolean;
  placeholder?: string;
  className?: string;
  description?: string;
  fieldType?: string;        // 'lookup' => commit under `${id}Id` as numbers
  multiselect?: boolean;     // v9 prop
  submitting?: boolean;      // matches your TextArea pattern
}

const REQUIRED_MSG = 'This is a required field and cannot be blank!';

const toKey = (k: unknown): string => (k == null ? '' : String(k));

const normalizeToStringArray = (input: unknown): string[] => {
  if (input == null) return [];
  if (typeof input === 'object' && (input as any).results && Array.isArray((input as any).results)) {
    return ((input as any).results as unknown[]).map(toKey);
  }
  if (Array.isArray(input)) return (input as unknown[]).map(toKey);
  return [toKey(input)];
};

export default function DropdownComponent(props: DropdownProps): JSX.Element {
  const {
    id,
    displayName,
    options,
    starterValue,
    isRequired: isRequiredProp = false,
    placeholder,
    className,
    description,
    fieldType,
    multiselect = false,
    submitting = false,
  } = props;

  const isLookup = fieldType === 'lookup';

  // ---- match TextArea state names / layout ----
  const [localVal, setLocalVal] = React.useState<string>('');        // used for disabled display text
  const [error, setError] = React.useState<string>('');
  const [isDisabled, setIsDisabled] = React.useState<boolean>(false);
  const [isHidden, setIsHidden] = React.useState<boolean>(false);

  const {
    FormData, GlobalFormData, GlobalErrorHandle,
    FormMode, AllDisableFields, AllHiddenFields,
    userBasedPerms, curUserInfo, listCols
  } = React.useContext(DynamicFormContext);

  // keep selection as strings for v9 selectedOptions
  const [selected, setSelected] = React.useState<string[]>([]);

  // map key->label for display
  const keyToText = React.useMemo(() => {
    const m = new Map<string, string>();
    for (const o of options) m.set(toKey(o.key), o.text);
    return m;
  }, [options]);

  const displayFromSelected = React.useCallback(
    (arr: string[]) => arr.map(k => keyToText.get(k) ?? k).join('; '),
    [keyToText]
  );

  const commitToContext = React.useCallback((arr: string[]) => {
    const targetId = isLookup ? `${id}Id` : id;

    if (arr.length === 0) {
      // SharePoint expects null when empty
      // eslint-disable-next-line @rushstack/no-new-null
      GlobalFormData(targetId, null);
      GlobalErrorHandle(targetId, props.isRequired ? REQUIRED_MSG : null);
      return;
    }

    if (isLookup) {
      const nums = arr.map(k => Number(k)).filter(n => Number.isFinite(n));
      // eslint-disable-next-line @rushstack/no-new-null
      GlobalFormData(targetId, multiselect ? nums : (nums[0] ?? null));
    } else {
      // eslint-disable-next-line @rushstack/no-new-null
      GlobalFormData(targetId, multiselect ? arr : (arr[0] ?? null));
    }

    GlobalErrorHandle(targetId, null);
  }, [GlobalFormData, GlobalErrorHandle, id, isLookup, multiselect, props.isRequired]);

  // ---- handleBlur (mirrors TextArea) ----
  const handleBlur = (): void => {
    if (props.isRequired === true && selected.length === 0) {
      setError(REQUIRED_MSG);
      GlobalErrorHandle(isLookup ? `${id}Id` : id, REQUIRED_MSG);
    } else {
      const joined = displayFromSelected(selected);
      setLocalVal(joined);
      setError('');
      GlobalErrorHandle(isLookup ? `${id}Id` : id, null);
      commitToContext(selected);
    }
  };

  // ---- handleChange (mirrors TextArea) ----
  const handleChange = (_e: unknown, data: DropdownOnOptionSelectData): void => {
    const next = (data.selectedOptions ?? []).map(toKey);
    setSelected(next);
    setLocalVal(displayFromSelected(next));
  };

  // ---- Initial render + edit/view prefill (mirrors TextArea) ----
  React.useEffect(() => {
    // EditForm or ViewForm
    if (FormMode === 4 || FormMode === 6) {
      const fldInternalName = isLookup ? `${id}Id` : id;
      if (FormData !== undefined) {
        const fieldValue = (FormData as any)[fldInternalName];
        const arr = normalizeToStringArray(fieldValue);
        setSelected(arr);
        setLocalVal(displayFromSelected(arr));
      }
    } else {
      // New form; seed from starterValue if provided
      const init = starterValue == null
        ? []
        : Array.isArray(starterValue)
          ? (starterValue as (string | number)[]).map(toKey)
          : [toKey(starterValue)];
      setSelected(init);
      setLocalVal(displayFromSelected(init));
    }

    // Disable or enable the field (same pattern as your TextArea)
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
      } as any;

      const results = formFieldsSetup(formFieldProps) || [];
      if (results.length > 0) {
        for (let i = 0; i < results.length; i += 1) {
          if (results[i].isDisabled !== undefined) setIsDisabled(results[i].isDisabled);
          if (results[i].isHidden !== undefined) setIsHidden(results[i].isHidden);
        }
      }
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  // ---- Submit disable hook (identical to your TextArea) ----
  React.useEffect(() => {
    if (props.submitting === true) {
      setIsDisabled(true);
    }
  }, [props.submitting]);

  const joined = localVal; // already semicolon-joined
  const disabledClass = isDisabled ? 'is-disabled' : '';
  const rootClassName = [className, disabledClass].filter(Boolean).join(' ');

  return (
    <div style={{ display: isHidden ? 'none' : 'block' }}>
      <Field
        label={props.displayName}
        {...(props.isRequired && { required: true })}
        validationMessage={error}
        validationState={error ? 'error' : undefined}
      >
        {isDisabled ? (
          // use disabled Input to keep gray-out and keep value visible
          <Input
            id={props.id}
            disabled
            value={joined}
            placeholder={placeholder || ''}
            className={rootClassName}
          />
        ) : (
          <Dropdown
            id={props.id}
            placeholder={placeholder}
            multiselect={!!multiselect}
            inlinePopup
            className={rootClassName}
            selectedOptions={selected}
            value={joined}
            onOptionSelect={handleChange}
            onBlur={handleBlur}
          >
            {options.map(o => (
              <Option key={toKey(o.key)} value={toKey(o.key)}>
                {o.text}
              </Option>
            ))}
          </Dropdown>
        )}

        {props.description !== '' && (
          <div className="descriptionText">{props.description}</div>
        )}
      </Field>
    </div>
  );
}






