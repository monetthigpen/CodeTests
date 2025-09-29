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
  multiSelect?: boolean;    // v8 prop
  multiselect?: boolean;    // v9 prop
  fieldType?: string;       // 'lookup' => commit under `${id}Id` as numbers
  options: { key: string | number; text: string }[];
  className?: string;
  description?: string;
  disabled?: boolean;
  submitting?: boolean;
}

const REQUIRED_MSG = 'This is a required field and cannot be blank!';

const toKey = (k: unknown): string => (k == null ? '' : String(k));

function normalizeToStringArray(input: unknown): string[] {
  if (!input) return [];

  if ((Array.isArray((input as any)?.results))) {
    return ((input as any).results as unknown[]).map(toKey);
  }

  if (Array.isArray(input)) {
    return (input as unknown[]).map(toKey);
  }

  if (typeof input === 'string' && input.includes(';')) {
    return input
      .split(';')
      .map(s => toKey(s.trim()))
      .filter(Boolean);
  }

  return [toKey(input)];
}

function clampToExisting(
  values: string[],
  opts: { key: string | number }[]
): string[] {
  const allowed = new Set(opts.map(o => toKey(o.key)));
  return values.filter(v => allowed.has(v));
}

export default function DropdownComponent(props: DropdownProps): JSX.Element {
  const {
    id,
    starterValue,
    displayName,
    isRequired: requiredProp = false,
    placeholder,
    multiSelect = false,
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

  // Controlled selection
  const [selectedOptions, setSelectedOptions] = React.useState<string[]>([]);
  const [error, setError] = React.useState<string>('');
  const [touched, setTouched] = React.useState<boolean>(false);

  // Lock/Cache for display text when disabled
  const [displayOverride, setDisplayOverride] = React.useState<string>('');
  const isLockedRef = React.useRef<boolean>(false);
  const didInitRef = React.useRef<boolean>(false);

  const keyToText = React.useMemo(() => {
    const m = new Map<string, string>();
    for (const o of options) m.set(toKey(o.key), o.text);
    return m;
  }, [options]);

  const reportError = React.useCallback(
    (msg: string) => {
      const targetId = isLookup ? `${id}LookupId` : id;
      setError(msg || '');
      GlobalErrorHandle(targetId, msg || null);
    },
    [GlobalErrorHandle, id, isLookup]
  );

  React.useEffect(() => {
    setIsRequired(!!requiredProp);
    setIsDisabled(!!disabledProp);
  }, [requiredProp, disabledProp]);

  // Submitting disables and locks display text
  React.useEffect(() => {
    if (submitting) {
      setIsDisabled(true);
      isLockedRef.current = true;
      const labels = selectedOptions.map(k => keyToText.get(k) ?? k);
      setDisplayOverride(labels.join('; '));
    }
  }, [submitting, selectedOptions, keyToText]);

  const ensureInOptions = (vals: string[]) => clampToExisting(vals, options);

  if (!isLockedRef.current) {
    if (!didInitRef.current) {
      if (FormMode !== 3) {
        const initArr =
          Array.isArray(starterValue)
            ? starterValue.map(toKey)
            : [toKey(starterValue)];
        setSelectedOptions(ensureInOptions(initArr));
      }
      didInitRef.current = true;
    } else {
      let raw;
      if (isLookup) {
        if (multiSelect) {
          const mLookup = (FormData as any)[`${id}`];
          if (mLookup?.length > 0) {
            raw = mLookup.map((v: any) => v.LookupId);
          } else {
            raw = [];
          }
        } else {
          raw = (FormData as any)[`${id}LookupId`];
        }
      } else {
        raw = (FormData as any)[id];
      }
      const arr = ensureInOptions(normalizeToStringArray(raw));
      setSelectedOptions(arr);
    }
  } else {
    const clamped = ensureInOptions(selectedOptions);
    if (clamped.length !== selectedOptions.length) {
      setSelectedOptions(clamped);
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
      HiddenList: AllHiddenFields,
      UserBasedList: userBasedPerms,
      curUserList: curUserInfo,
      curField: displayName,
      formStateData: FormData,
      ListColumns: listCols,
    } as any;

    const results = formFieldsSetup(formFieldProps) || [];
    if (results.length > 0) {
      for (let i = 0; i < results.length; i++) {
        if (results[i].isDisabled !== undefined) setIsDisabled(results[i].isDisabled);
        if (results[i].isHidden !== undefined) setIsHidden(results[i].isHidden);
      }
    }
    if (!isLockedRef.current && isDisabled) {
      const labels = selectedOptions.map(k => keyToText.get(k) ?? k);
      setDisplayOverride(labels.join('; '));
    }
  }

  reportError('');
  setTouched(false);

  const validate = React.useCallback((): string => {
    return (isRequired && selectedOptions.length === 0) ? REQUIRED_MSG : '';
  }, [isRequired, selectedOptions]);

  // Commit: send null when empty; numbers for lookup
  const commitValue = React.useCallback(() => {
    const err = validate();
    reportError(err);

    const targetId = isLookup ? `${id}Id` : id;

    if (isLookup) {
      const nums = selectedOptions.map(k => Number(k)).filter(n => Number.isFinite(n));
      GlobalFormData[targetId] = nums.length === 0 ? null : multiSelect ? nums : nums[0];
    } else {
      GlobalFormData[targetId] = selectedOptions.length === 0 ? null : multiSelect ? selectedOptions : selectedOptions[0];
    }

    const labels = selectedOptions.map(k => keyToText.get(k) ?? k);
    setDisplayOverride(labels.join('; '));
  }, [validate, reportError, GlobalFormData, id, isLookup, multiSelect, selectedOptions, keyToText]);

  const handleOptionSelect = (
    e: unknown,
    data: { optionValue: string | number; selectedOptions: (string | number)[] }
  ) => {
    const next = (data.selectedOptions ?? []).map(toKey);
    setSelectedOptions(next);
    if (!touched) reportError(isRequired && next.length === 0 ? REQUIRED_MSG : '');
  };

  const handleBlur = () => {
    setTouched(true);
    commitValue();
  };

  // SemiColon-joined labels for display
  const selectedLabels = selectedOptions.map(k => keyToText.get(k) ?? k);
  const joinedText = selectedLabels.join('; ');
  const triggerText = displayOverride || joinedText;
  const triggerPlaceholder = triggerText || placeholder || '';

  // Build class and attributes so parent CSS gray-out continues to work
  const disabledClass = isDisabled ? 'is-disabled' : '';
  const rootClassName = [className, disabledClass].filter(Boolean).join(' ');

  return (
    <div
      style={{ display: isHidden ? 'none' : 'block' }}
    >
      <Field
        label={displayName}
        required={isRequired}
        validationMessage={error ? error : undefined}
        validationState={error ? 'error' : undefined}
      >
        {isDisabled ? (
          // Disabled Input to retain gray-out visuals and keep text visible
          <Input
            id={id}
            disabled
            value={triggerText}
            placeholder={triggerPlaceholder}
            className={rootClassName}
            aria-disabled="true"
            data-disabled="true"
          />
        ) : (
          <Dropdown
            id={id}
            multiselect={multiSelect}
            disabled={false}
            inlinePopup={true}
            selectedOptions={selectedOptions}
            onOptionSelect={handleOptionSelect}
            onBlur={handleBlur}
            className={rootClassName}
            value={triggerText}
            placeholder={triggerPlaceholder}
            title={triggerText || displayName}
            aria-label={triggerText || displayName}
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