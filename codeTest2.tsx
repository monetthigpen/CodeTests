import * as React from 'react';
import { Field, Dropdown, Option } from '@fluentui/react-components';
import { DynamicFormContext } from './DynamicFormContext';
import formFieldsSetup, { FormFieldsProps } from './formFieldBased';

/* ------------------------------ Props ------------------------------ */

interface DropdownProps {
  id: string;
  starterValue?: string | number | Array<string | number>;
  displayName: string;
  isRequired?: boolean;
  placeholder?: string;
  multiselect?: boolean;                 // v9 prop
  fieldType?: string;                    // 'lookup' => commit under `${id}Id` as numbers
  options: { key: string | number; text: string }[];
  className?: string;
  description?: string;
  disabled?: boolean;
  submitting?: boolean;
}

const REQUIRED_MSG = 'This is a required field and cannot be blank!';

/* ------------------------------ Helpers ------------------------------ */

const toKey = (k: unknown): string => (k == null ? '' : String(k));

function normalizeToStringArray(input: unknown): string[] {
  if (input == null) return [];
  if (Array.isArray((input as any)?.results)) {
    return ((input as any).results as unknown[]).map(toKey);
  }
  if (Array.isArray(input)) {
    return (input as unknown[]).map(toKey);
  }
  if (typeof input === 'string' && input.includes(';')) {
    return input.split(';').map(s => toKey(s.trim())).filter(Boolean);
  }
  return [toKey(input)];
}

function clampToExisting(values: string[], opts: { key: string | number }[]): string[] {
  const allowed = new Set(opts.map(o => toKey(o.key)));
  return values.filter(v => allowed.has(v));
}

/* ------------------------------ Component ------------------------------ */

export default function DropdownComponent(props: DropdownProps): JSX.Element {
  const {
    id,
    starterValue,
    displayName,
    isRequired: requiredProp = false,
    placeholder,
    multiselect = false, // v9 prop
    fieldType,
    options,
    className,
    description,
    disabled: disabledProp = false,
    submitting = false,
  } = props;

  const isLookup = fieldType === 'lookup';

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

  const [isRequired, setIsRequired] = React.useState<boolean>(!!requiredProp);
  const [isDisabled, setIsDisabled] = React.useState<boolean>(!!disabledProp);
  const [isHidden, setIsHidden] = React.useState<boolean>(false);

  const [selectedOptions, setSelectedOptions] = React.useState<string[]>([]);
  const [error, setError] = React.useState<string>('');
  const [touched, setTouched] = React.useState<boolean>(false);

  // Mirror error -> GlobalErrorHandle (null when empty)
  const reportError = React.useCallback(
    (msg: string) => {
      const targetId = isLookup ? `${id}Id` : id;
      setError(msg || '');
      GlobalErrorHandle(targetId, msg || null);
    },
    [GlobalErrorHandle, id, isLookup]
  );

  React.useEffect(() => {
    setIsRequired(!!requiredProp);
    setIsDisabled(!!disabledProp);
  }, [requiredProp, disabledProp]);

  // Submitting disables the field
  React.useEffect(() => {
    if (submitting) setIsDisabled(true);
  }, [submitting]);

  // Prefill + rules + display mode
  React.useEffect(() => {
    const ensureInOptions = (vals: string[]) => clampToExisting(vals, options);

    if (FormMode === 8) {
      const initArr = starterValue
        ? Array.isArray(starterValue)
          ? starterValue.map(toKey)
          : [toKey(starterValue)]
        : [];
      setSelectedOptions(ensureInOptions(initArr));
    } else {
      const raw = FormData
        ? (isLookup ? (FormData as any)[`${id}Id`] : (FormData as any)[id])
        : undefined;
      const arr = ensureInOptions(normalizeToStringArray(raw));
      setSelectedOptions(arr);
    }

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
        for (let i = 0; i < results.length; i++) {
          if (results[i].isDisabled !== undefined) setIsDisabled(results[i].isDisabled);
          if (results[i].isHidden !== undefined) setIsHidden(results[i].isHidden);
        }
      }
    }

    reportError('');
    setTouched(false);
  }, [
    FormData,
    FormMode,
    starterValue,
    options,
    isLookup,
    id,
    displayName,
    AllDisableFields,
    AllHiddenFields,
    userBasedPerms,
    curUserInfo,
    listCols,
    reportError,
  ]);

  // Validate and commit value
  const validate = React.useCallback((): string => {
    if (isRequired && selectedOptions.length === 0) return REQUIRED_MSG;
    return '';
  }, [isRequired, selectedOptions]);

  const commitValue = React.useCallback(() => {
    const err = validate();
    reportError(err);

    const targetId = isLookup ? `${id}Id` : id;
    if (isLookup) {
      const nums = selectedOptions
        .map(k => Number(k))
        .filter(n => Number.isFinite(n));
      GlobalFormData(targetId, multiselect ? nums : nums[0] ?? null);
    } else {
      GlobalFormData(targetId, multiselect ? selectedOptions : selectedOptions[0] ?? null);
    }
  }, [validate, reportError, GlobalFormData, id, isLookup, multiselect, selectedOptions]);

  const handleOptionSelect = (
    _e: unknown,
    data: { optionValue?: string | number; selectedOptions: (string | number)[] }
  ) => {
    const next = (data.selectedOptions ?? []).map(toKey);
    setSelectedOptions(next);
    if (touched) reportError(isRequired && next.length === 0 ? REQUIRED_MSG : '');
  };

  const handleBlur = () => {
    setTouched(true);
    commitValue();
  };

  const displayText = selectedOptions.length ? selectedOptions.join('; ') : '';
  const effectivePlaceholder = displayText || placeholder || '';
  const hasError = !!error;

  return (
    <div style={{ display: isHidden ? 'none' : 'block' }}>
      <Field
        label={displayName}
        required={isRequired}
        validationMessage={hasError ? error : undefined}
        validationState={hasError ? 'error' : undefined}
      >
        <Dropdown
          id={id}
          placeholder={effectivePlaceholder}
          multiselect={multiselect}
          disabled={isDisabled}
          inlinePopup={true}
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

        {description && (
          <div className="descriptionText">{description}</div>
        )}
      </Field>
    </div>
  );
}
