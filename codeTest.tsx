import * as React from 'react';
import { Field, Dropdown, Option, Input } from '@fluentui/react-components';
import { DynamicFormContext } from './DynamicFormContext';
import formFieldsSetup, { FormFieldsProps } from './formFieldBased';

// Derive types from the component (works across library versions)
type OnOptionSelect = NonNullable<React.ComponentProps<typeof Dropdown>['onOptionSelect']>;
type OnOptionSelectEvent = Parameters<OnOptionSelect>[0];
type OnOptionSelectData = Parameters<OnOptionSelect>[1];

type OptionShape = { key: string | number; text: string };

interface DropdownProps {
  id: string;
  displayName: string;
  options: OptionShape[];
  starterValue?: string | number | Array<string | number>;
  isRequired?: boolean;
  placeholder?: string;
  className?: string;
  description?: string;
  fieldType?: string;        // 'lookup' => commit under `${id}Id` as numbers
  multiselect?: boolean;     // v9 prop
  multiSelect?: boolean;     // alias to match older usage
  disabled?: boolean;
  submitting?: boolean;
}

const REQUIRED_MSG = 'This is a required field and cannot be blank!';

const toKey = (k: unknown): string => (k == null ? '' : String(k));

const normalizeToStringArray = (input: unknown): string[] => {
  if (input == null) return [];
  if (typeof input === 'object' && input !== null && Array.isArray((input as { results?: unknown[] }).results)) {
    return ((input as { results: unknown[] }).results).map(toKey);
  }
  if (Array.isArray(input)) return (input as unknown[]).map(toKey);
  if (typeof input === 'string' && input.includes(';')) {
    return input.split(';').map(s => toKey(s.trim())).filter(Boolean);
  }
  return [toKey(input)];
};

const clampToExisting = (values: string[], opts: OptionShape[]): string[] => {
  const allowed = new Set(opts.map(o => toKey(o.key)));
  return values.filter(v => allowed.has(v));
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
    multiselect,
    multiSelect,
    disabled: disabledProp = false,
    submitting = false,
  } = props;

  const isLookup = fieldType === 'lookup';
  const isMulti = !!(multiselect ?? multiSelect);

  // Match the state naming style in your examples
  const [localVal, setLocalVal] = React.useState<string>('');      // visible text (semicolon-joined)
  const [error, setError] = React.useState<string>('');
  const [isDisabled, setIsDisabled] = React.useState<boolean>(!!disabledProp);
  const [isHidden, setIsHidden] = React.useState<boolean>(false);

  // Keep internal selections as keys (strings) to drive Dropdown
  const [selectedKeys, setSelectedKeys] = React.useState<string[]>([]);

  const {
    FormData,
    GlobalFormData,
    GlobalErrorHandle,
    FormMode,
    AllDisableFields,
    AllHiddenFields,
    userBasedPerms,
    curUserInfo,
    listCols,
  } = React.useContext(DynamicFormContext);

  // Map key -> label for display
  const keyToText = React.useMemo(() => {
    const m = new Map<string, string>();
    for (const o of options) m.set(toKey(o.key), o.text);
    return m;
  }, [options]);

  const textFromKeys = React.useCallback(
    (arr: string[]): string => arr.map(k => keyToText.get(k) ?? k).join(';'),
    [keyToText]
  );

  // -------------------- useEffect #1: Initial render (prefill + rules) --------------------
  React.useEffect(() => {
    // Prefill value
    let initKeys: string[] = [];
    if (FormMode === 8) {
      if (starterValue != null) {
        initKeys = Array.isArray(starterValue)
          ? (starterValue as (string | number)[]).map(toKey)
          : [toKey(starterValue)];
      }
    } else {
      const raw = (FormData as any)
        ? (isLookup ? (FormData as any)[`${id}Id`] : (FormData as any)[id])
        : undefined;
      initKeys = normalizeToStringArray(raw);
    }
    initKeys = clampToExisting(initKeys, options);
    setSelectedKeys(initKeys);
    setLocalVal(textFromKeys(initKeys));

    // Disable/Hide rules
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
      for (const r of results) {
        if (r.isDisabled !== undefined) setIsDisabled(!!r.isDisabled);
        if (r.isHidden !== undefined) setIsHidden(!!r.isHidden);
      }
    }

    // Reset error on init
    setError('');
    GlobalErrorHandle(isLookup ? `${id}Id` : id, null);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  // -------------------- useEffect #2: Submitting → disable (and keep display) --------------------
  React.useEffect(() => {
    if (submitting === true) {
      setIsDisabled(true);
      setLocalVal(textFromKeys(selectedKeys)); // keep visible text after disable
    }
  }, [submitting, selectedKeys, textFromKeys]);

  // -------------------- Handlers --------------------
  const handleChange = (_e: OnOptionSelectEvent, data: OnOptionSelectData): void => {
    const next = (data.selectedOptions ?? []).map(v => String(v));
    setSelectedKeys(next);
    setLocalVal(textFromKeys(next));
  };

  const handleBlur = (): void => {
    // Validate
    if ((isRequiredProp || props.isRequired) && selectedKeys.length === 0) {
      setError(REQUIRED_MSG);
      GlobalErrorHandle(isLookup ? `${id}Id` : id, REQUIRED_MSG);
      // eslint-disable-next-line @rushstack/no-new-null
      GlobalFormData(isLookup ? `${id}Id` : id, null);
      return;
    }

    setError('');
    GlobalErrorHandle(isLookup ? `${id}Id` : id, null);

    // Commit
    if (isLookup) {
      const nums = selectedKeys.map(k => Number(k)).filter(n => Number.isFinite(n));
      // eslint-disable-next-line @rushstack/no-new-null
      GlobalFormData(isMulti ? `${id}Id` : `${id}Id`, nums.length === 0 ? null : (isMulti ? nums : nums[0]));
    } else {
      // eslint-disable-next-line @rushstack/no-new-null
      GlobalFormData(id, selectedKeys.length === 0 ? null : (isMulti ? selectedKeys : selectedKeys[0]));
    }
  };

  // -------------------- Render --------------------
  const effectivePlaceholder = localVal || placeholder || '';
  const disabledClass = isDisabled ? 'is-disabled' : '';
  const rootClassName = [className, disabledClass].filter(Boolean).join(' ');

  return (
    <div style={{ display: isHidden ? 'none' : 'block' }}>
      <Field
        label={displayName}
        {...(isRequiredProp && { required: true })}
        validationMessage={error}
        validationState={error ? 'error' : undefined}
      >
        {isDisabled ? (
          <Input
            id={id}
            type="text"
            value={localVal}
            disabled
            className={rootClassName}
            placeholder={effectivePlaceholder}
          />
        ) : (
          <Dropdown
            id={id}
            multiselect={isMulti}
            inlinePopup
            value={localVal}
            placeholder={effectivePlaceholder}
            selectedOptions={selectedKeys}
            onOptionSelect={handleChange}
            onBlur={handleBlur}
            className={rootClassName}
          >
            {options.map(o => (
              <Option key={toKey(o.key)} value={toKey(o.key)}>
                {o.text}
              </Option>
            ))}
          </Dropdown>
        )}

        {description !== '' && description !== undefined && (
          <div className="descriptionText">{description}</div>
        )}
      </Field>
    </div>
  );
}







