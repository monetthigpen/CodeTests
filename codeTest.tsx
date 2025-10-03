import * as React from 'react';
import { Field, Dropdown, Option, Input } from '@fluentui/react-components';
import { DynamicFormContext } from './DynamicFormContext';
import { formFieldsSetup, FormFieldsProps } from './formFieldBased';

interface DropdownProps {
  id: string;
  starterValue?: string | number | Array<string | number>;
  displayName: string;
  isRequired?: boolean;
  placeholder?: string;
  multiSelect?: boolean;   // v8 prop
  multiselect?: boolean;   // v9 prop
  fieldType?: string;      // 'lookup' => commit under `${id}Id` as numbers
  options: { key: string | number; text: string }[];
  className?: string;
  description?: string;
  disabled?: boolean;
  submitting?: boolean;
}

const REQUIRED_MSG = 'This is a required field and cannot be blank!';

const toKey = (k: unknown): string => (k == null ? '' : String(k));

function normalizeToStringArray(input: unknown): string[] {
  if (input == null) return [];
  if ((input as any)?.results) {
    return ((input as any).results as unknown[]).map(toKey);
  }
  if (Array.isArray(input)) return (input as unknown[]).map(toKey);
  if (typeof input === 'string') return input.split('.').map(s => toKey(s.trim())).filter(Boolean);
  return [toKey(input)];
}

function clampToExisting(values: string[], opts: { key: string | number }[]): string[] {
  const allowed = new Set(opts.map(o => toKey(o.key)));
  return values.filter(v => allowed.has(v));
}

export default function DropdownComponent(props: DropdownProps): JSX.Element {
  const elemRef = React.useRef<HTMLDivElement | null>(null);

  const {
    id,
    starterValue,
    displayName,
    isRequired: requiredProp = false,
    placeholder,
    multiSelect = false,
    multiselect = false,
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
    GlobalRefs,
    AllDisableFields,
    AllHiddenFields,
    userBasedPerms,
    curUserInfo,
    listCols,
  } = React.useContext(DynamicFormContext);

  const [isRequired, setIsRequired] = React.useState<boolean>(!!requiredProp);
  const [isDisabled, setIsDisabled] = React.useState<boolean>(!!disabledProp);
  const [isHidden, setIsHidden] = React.useState<boolean>(false);
  const [defaultDisable, setDefaultDisable] = React.useState<boolean>(false);

  const [selectedOptions, setSelectedOptions] = React.useState<string[]>([]);
  const [error, setError] = React.useState<string>('');
  const [touched, setTouched] = React.useState<boolean>(false);

  const [displayOverride, setDisplayOverride] = React.useState<string>('');
  const isLockedRef = React.useRef<boolean>(false);
  const didInitRef = React.useRef<boolean>(false);

  const keyToText = React.useMemo(() => {
    const m = new Map<string, string>();
    for (const o of options) m.set(toKey(o.key), o.text);
    return m;
  }, [options]);

  const reportError = React.useCallback((msg: string) => {
    const targetId = isLookup ? `${id}Id` : id;
    setError(msg || '');
    GlobalErrorHandle(targetId, msg || null);
  }, [GlobalErrorHandle, id, isLookup]);

  React.useEffect(() => {
    setIsRequired(!!requiredProp);
    setIsDisabled(!!disabledProp);
  }, [requiredProp, disabledProp]);

  // Prefill and disable
  React.useEffect(() => {
    const ensureInOptions = (vals: string[]) => clampToExisting(vals, options);

    if (!isLockedRef.current) {
      if (!didInitRef.current) {
        if (FormMode === 3) {
          const initArr = Array.isArray(starterValue)
            ? starterValue.map(toKey)
            : starterValue != null
              ? [toKey(starterValue)]
              : [];
          setSelectedOptions(ensureInOptions(initArr));
        } else {
          let raw: any;
          if (isLookup) {
            const lookup = (FormData as any)?.[`${id}Id`];
            raw = lookup ? lookup.map((v: any) => v.LookupId) : [];
          } else {
            raw = (FormData as any)?.[id];
          }
          const arr = ensureInOptions(normalizeToStringArray(raw));
          setSelectedOptions(arr);
        }
        didInitRef.current = true;
      } else {
        const clamped = ensureInOptions(selectedOptions);
        if (clamped.length !== selectedOptions.length) {
          setSelectedOptions(clamped);
        }
      }
    }

    if (FormMode === 4) {
      setIsDisabled(true);
      isLockedRef.current = true;
      const labels = selectedOptions.map(k => keyToText.get(k) ?? k);
      setDisplayOverride(labels.join('; '));
    } else {
      const formFieldProps: FormFieldsProps = {
        disableList: AllDisableFields,
        hiddenList: AllHiddenFields,
        userBasedList: userBasedPerms,
        curUserList: curUserInfo,
        curField: displayName,
        formStateData: FormData,
        listColumns: (Array.isArray(listCols) ? listCols : []) as string[],
      };
      const results = formFieldsSetup(formFieldProps) ?? [];

      if (results.length > 0) {
        for (let i = 0; i < results.length; i++) {
          if (typeof results[i].isDisabled === 'boolean') setIsDisabled(results[i].isDisabled ?? false);
          if (typeof results[i].isHidden === 'boolean') setIsHidden(results[i].isHidden ?? false);
          setDefaultDisable(results[i].isDisabled ?? false);
        }
      }
    }
  }, [FormData, FormMode, id, displayName, options, isLookup, starterValue,
      AllDisableFields, AllHiddenFields, userBasedPerms, curUserInfo, listCols]);

  const validate = React.useCallback((): string => {
    if (isRequired && selectedOptions.length === 0) return REQUIRED_MSG;
    return '';
  }, [isRequired, selectedOptions]);

  const commitValue = React.useCallback(() => {
    const err = validate();
    reportError(err);

    const targetId = isLookup ? `${id}Id` : id;

    if (isMulti(multiSelect, multiselect)) {
      const nums = selectedOptions.map(k => Number(k)).filter(n => Number.isFinite(n));
      GlobalFormData(targetId, nums.length === 0 ? null : nums);
    } else {
      GlobalFormData(targetId, selectedOptions.length === 0 ? null : selectedOptions[0]);
    }
  }, [validate, reportError, GlobalFormData, id, isLookup, selectedOptions, multiSelect, multiselect]);

  const handleOptionSelect = (
    event: React.SyntheticEvent,
    data: { optionValue: string | number; selectedOptions: (string | number)[] }
  ) => {
    const next = (data.selectedOptions ?? []).map(toKey);
    setSelectedOptions(next);
    setTouched(true);
    commitValue();
  };

  const handleBlur = () => {
    setTouched(true);
    commitValue();
  };

  // Helpers for joined labels
  const selectedLabels = selectedOptions.map(k => keyToText.get(k) ?? k);
  const joinedText = selectedLabels.join('; ');
  const visibleText = displayOverride || joinedText;
  const triggerText = visibleText || '';
  const triggerPlaceholder = triggerText || placeholder || '';
  const hasError = !!error;

  if (isHidden) return null;

  return (
    <div style={{ display: isHidden ? 'none' : 'block' }}
         className={[className, isDisabled ? 'is-disabled' : ''].filter(Boolean).join(' ')}
         data-disabled={isDisabled ? 'true' : undefined}
         aria-disabled={isDisabled ? 'true' : undefined}>

      <Field
        label={displayName}
        required={isRequired}
        validationMessage={hasError ? error : undefined}
        validationState={hasError ? 'error' : undefined}
      >
        {isDisabled ? (
          <Input
            id={id}
            disabled
            value={triggerText}
            placeholder={triggerPlaceholder}
            className={className}
            aria-disabled="true"
            data-disabled="true"
          />
        ) : (
          <Dropdown
            id={id}
            multiselect={!!multiSelect || !!multiselect}
            disabled={false}
            selectedOptions={selectedOptions}
            onOptionSelect={handleOptionSelect}
            onBlur={handleBlur}
            className={className}
            value={triggerText}
            placeholder={triggerPlaceholder}
            title={triggerText}
            aria-label={triggerText || displayName}
            ref={elemRef as unknown as React.RefObject<HTMLButtonElement>}
          >
            {options.map(o => (
              <Option key={toKey(o.key)} value={toKey(o.key)}>
                {o.text}
              </Option>
            ))}
          </Dropdown>
        )}
      </Field>
      {description && <div className="descriptionText">{description}</div>}
    </div>
  );
}

// helper to unify multiSelect props
function isMulti(multiSelect?: boolean, multiselect?: boolean): boolean {
  return !!multiSelect || !!multiselect;
}




