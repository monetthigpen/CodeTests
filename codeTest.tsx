import * as React from 'react';
import { Field, Dropdown, Option, Input } from '@fluentui/react-components';
import { DynamicFormContext } from './DynamicFormContext';
import formFieldsSetup, { FormFieldsProps } from './formFieldBased';

// Derive handler types directly from the Dropdown component
type OnOptionSelect = NonNullable<React.ComponentProps<typeof Dropdown>['onOptionSelect']>;
type OnOptionSelectEvent = Parameters<OnOptionSelect>[0];
type OnOptionSelectData = Parameters<OnOptionSelect>[1];

type OptionShape = { key: string | number; text: string };

interface DropdownProps {
  id: string;
  displayName: string;
  options: OptionShape[];

  // optional inputs
  starterValue?: string | number | Array<string | number>;
  isRequired?: boolean;
  placeholder?: string;
  className?: string;
  description?: string;
  fieldType?: string;          // 'lookup' -> commit under `${id}Id` as numbers
  multiselect?: boolean;       // v9 prop
  multiSelect?: boolean;       // alias (older usage)
  disabled?: boolean;
  submitting?: boolean;
  locale?: string;             // parent passes this in some usages (avoid JSX type error)
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

  // match your other components: localVal, error, isDisabled, isHidden
  const [localVal, setLocalVal] = React.useState<string>('');      // semicolon-joined label text
  const [error, setError] = React.useState<string>('');
  const [isDisabled, setIsDisabled] = React.useState<boolean>(!!disabled);
  const [isHidden, setIsHidden] = React.useState<boolean>(false);

  // drive Dropdown with selected keys (strings)
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

  // -------------------- useEffect #1: Initial prefill + rules (runs once) --------------------
  React.useEffect(() => {
    // prefill selection
    let initKeys: string[] = [];
    if (FormMode === 8) {
      initKeys =
        starterValue == null
          ? []
          : Array.isArray(starterValue)
            ? (starterValue as (string | number)[]).map(toKey)
            : [toKey(starterValue)];
    } else {
      const raw = (FormData as any)
        ? (isLookup ? (FormData as any)[`${id}Id`] : (FormData as any)[id])
        : undefined;
      initKeys = normalizeToStringArray(raw);
    }
    initKeys = clampToExisting(initKeys, options);
    setSelectedKeys(initKeys);
    setLocalVal(textFromKeys(initKeys));

    // rules
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

    // reset error at init
    setError('');
    GlobalErrorHandle(isLookup ? `${id}Id` : id, null);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  // -------------------- useEffect #2: submitting -> disable & keep visible text --------------------
  React.useEffect(() => {
    if (submitting) {
      setIsDisabled(true);
      setLocalVal(textFromKeys(selectedKeys)); // keep value shown after disable
    }
  }, [submitting, selectedKeys, textFromKeys]);

  // -------------------- handlers --------------------
  const handleChange = (_e: OnOptionSelectEvent, data: OnOptionSelectData): void => {
    const next = (data.selectedOptions ?? []).map(v => String(v));
    setSelectedKeys(next);
    setLocalVal(textFromKeys(next));
  };

  const handleBlur = (): void => {
    // validate
    if (isRequired && selectedKeys.length === 0) {
      setError(REQUIRED_MSG);
      GlobalErrorHandle(isLookup ? `${id}Id` : id, REQUIRED_MSG);
      // eslint-disable-next-line @rushstack/no-new-null
      GlobalFormData(isLookup ? `${id}Id` : id, null);
      return;
    }

    setError('');
    GlobalErrorHandle(isLookup ? `${id}Id` : id, null);

    // commit value (null when empty, numbers for lookup)
    if (isLookup) {
      const nums = selectedKeys.map(k => Number(k)).filter((n): n is number => Number.isFinite(n));
      // eslint-disable-next-line @rushstack/no-new-null
      GlobalFormData(`${id}Id`, nums.length === 0 ? null : (isMulti ? nums : nums[0]));
    } else {
      // eslint-disable-next-line @rushstack/no-new-null
      GlobalFormData(id, selectedKeys.length === 0 ? null : (isMulti ? selectedKeys : selectedKeys[0]));
    }
  };

  // -------------------- render --------------------
  const effectivePlaceholder = localVal || placeholder || '';
  const disabledClass = isDisabled ? 'is-disabled' : '';
  const rootClassName = [className, disabledClass].filter(Boolean).join(' ');

  return (
    <div style={{ display: isHidden ? 'none' : 'block' }}>
      <Field
        label={displayName}
        {...(isRequired && { required: true })}
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
            inlinePopup={true}
            className={rootClassName}
            value={localVal}
            placeholder={effectivePlaceholder}
            selectedOptions={selectedKeys}
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

        {description ? <div className="descriptionText">{description}</div> : null}
      </Field>
    </div>
  );
}







