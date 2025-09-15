import * as React from 'react';
import { Field, Dropdown, Option } from '@fluentui/react-components';
import { DynamicFormContext } from './DynamicFormContext';
import formFieldsSetup, { FormFieldsProps } from './formFieldBased';

// Types derived from Dropdown (version-safe)
type OnOptionSelect = NonNullable<React.ComponentProps<typeof Dropdown>['onOptionSelect']>;
type OnOptionSelectEvent = Parameters<OnOptionSelect>[0];
type OnOptionSelectData = Parameters<OnOptionSelect>[1];

export interface DropdownProps {
  id: string;
  displayName: string;
  options: { key: string | number; text: string }[];
  starterValue?: string | number | Array<string | number>;
  isRequired?: boolean;
  placeholder?: string;
  className?: string;
  description?: string;
  fieldType?: string;      // 'lookup' => commit under `${id}LookupId` as numbers
  multiselect?: boolean;   // v9
  multiSelect?: boolean;   // v8 alias
  disabled?: boolean;
  submitting?: boolean;
}

const REQUIRED_MSG = 'This is a required field and cannot be blank!';
const LOOKUP_SUFFIX = 'LookupId';

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

const clampToExisting = (
  values: string[],
  opts: Array<{ key: string | number }>
): string[] => {
  const allowed = new Set(opts.map(o => toKey(o.key)));
  return values.filter(v => allowed.has(v));
};

export default function DropdownComponent(props: DropdownProps): JSX.Element {
  const {
    id,
    displayName,
    options,
    starterValue,
    isRequired = false,
    placeholder,
    className,
    description,
    fieldType,
    multiselect,
    multiSelect,
    disabled = false,
    submitting = false,
  } = props;

  const isLookup = fieldType === 'lookup';
  const isMulti = !!(multiselect ?? multiSelect);
  const targetId = isLookup ? `${id}${LOOKUP_SUFFIX}` : id;

  // Visual/selection state
  const [localVal, setLocalVal] = React.useState<string>('');      // semicolon-joined labels
  const [selectedKeys, setSelectedKeys] = React.useState<string[]>([]);
  const [error, setError] = React.useState<string>('');
  const [isDisabled, setIsDisabled] = React.useState<boolean>(!!disabled);
  const [isHidden, setIsHidden] = React.useState<boolean>(false);

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

  // key -> label
  const keyToText = React.useMemo(() => {
    const m = new Map<string, string>();
    for (const o of options) m.set(toKey(o.key), o.text);
    return m;
  }, [options]);

  const textFromKeys = React.useCallback(
    (arr: string[]): string => arr.map(k => keyToText.get(k) ?? k).join(';'),
    [keyToText]
  );

  // -------------------- useEffect #1: Initial prefill + rules (once) --------------------
  React.useEffect((): void => {
    // Prefill from starterValue (New) or FormData (Edit/View)
    let initKeys: string[] = [];
    if (FormMode === 8) {
      initKeys =
        starterValue == null
          ? []
          : Array.isArray(starterValue)
          ? (starterValue as (string | number)[]).map(toKey)
          : [toKey(starterValue)];
    } else {
      const formBag = (FormData ?? {}) as Record<string, unknown>; // ✅ safe narrowing
      const key = isLookup ? `${id}${LOOKUP_SUFFIX}` : id;
      const raw = formBag[key];
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
      } as unknown as FormFieldsProps;

      const results = formFieldsSetup(formFieldProps) || [];
      for (const r of results) {
        if (r.isDisabled !== undefined) setIsDisabled(!!r.isDisabled);
        if (r.isHidden !== undefined) setIsHidden(!!r.isHidden);
      }
    }

    // Clear error on init
    setError('');
    // eslint-disable-next-line @rushstack/no-new-null
    GlobalErrorHandle(targetId, null);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []); // once

  // -------------------- useEffect #2: submitting → disable & keep visible text --------------------
  React.useEffect((): void => {
    if (submitting) {
      setIsDisabled(true);
      setLocalVal(textFromKeys(selectedKeys)); // keep text visible after disabling
    }
  }, [submitting, selectedKeys, textFromKeys]);

  // -------------------- Handlers --------------------
  const onOptionSelect = (_e: OnOptionSelectEvent, data: OnOptionSelectData): void => {
    const next = (data.selectedOptions ?? []).map(v => String(v));
    setSelectedKeys(next);
    setLocalVal(textFromKeys(next));
  };

  const handleBlur = (): void => {
    // Validate
    if (isRequired && selectedKeys.length === 0) {
      setError(REQUIRED_MSG);
      GlobalErrorHandle(targetId, REQUIRED_MSG);
      // eslint-disable-next-line @rushstack/no-new-null
      GlobalFormData(targetId, null);
      return;
    }

    setError('');
    // eslint-disable-next-line @rushstack/no-new-null
    GlobalErrorHandle(targetId, null);

    // Commit value: null when empty; numbers for lookup
    if (isLookup) {
      const nums = selectedKeys.map(k => Number(k)).filter((n): n is number => Number.isFinite(n));
      // eslint-disable-next-line @rushstack/no-new-null
      GlobalFormData(targetId, nums.length === 0 ? null : (isMulti ? nums : nums[0]));
    } else {
      // eslint-disable-next-line @rushstack/no-new-null
      GlobalFormData(targetId, selectedKeys.length === 0 ? null : (isMulti ? selectedKeys : selectedKeys[0]));
    }
  };

  // -------------------- Render --------------------
  const effectiveClass = className ?? 'fieldClass';
  const effectivePlaceholder = localVal || placeholder || '';

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
          className={effectiveClass}
          multiselect={isMulti}
          inlinePopup={true}
          disabled={isDisabled}
          value={localVal}
          placeholder={effectivePlaceholder}
          selectedOptions={selectedKeys}
          onOptionSelect={onOptionSelect}
          onBlur={handleBlur}
          title={localVal}
          aria-label={localVal || displayName}
        >
          {options.map(o => (
            <Option key={toKey(o.key)} value={toKey(o.key)}>
              {o.text}
            </Option>
          ))}
        </Dropdown>

        {description ? <div className="descriptionText">{description}</div> : null}
      </Field>
    </div>
  );
}







