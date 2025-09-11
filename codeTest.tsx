import * as React from 'react';
import { Field, Dropdown, Option } from '@fluentui/react-components';
import { DynamicFormContext } from './DynamicFormContext';
import { formFieldsSetup, FormFieldsProps } from './formFieldBased';

type OptionItem = { key: string | number; text: string };

export interface DropdownFieldProps {
  id: string;
  displayName: string;
  starterValue?: string | number | Array<string | number>;
  isRequired?: boolean;
  disabled?: boolean;
  placeholder?: string;
  multiSelect?: boolean;   // v8-style name
  multiselect?: boolean;   // v9 prop name
  options: OptionItem[];
  fieldType?: string;      // "lookup" to send numeric Id(s)
  className?: string;
  description?: string;
  submitting?: boolean;    // drives disabled via its own effect
}

const REQUIRED_MSG = 'This is a required field and cannot be blank!';

// ---------- helpers ----------
const toKey = (k: unknown): string => (k == null ? '' : String(k));

function normalizeToStringArray(input: unknown): string[] {
  if (input == null) return [];

  if (Array.isArray((input as any)?.results)) {
    return ((input as any).results as unknown[]).map(toKey);
  }

  if (Array.isArray(input)) {
    const arr = input as unknown[];
    if (arr.length && typeof arr[0] === 'object' && arr[0] !== null) {
      return arr.map((o: any) => toKey(o?.Id ?? o?.id ?? o?.Key ?? o?.value ?? o));
    }
    return arr.map(toKey);
  }

  if (typeof input === 'string' && input.includes(';')) {
    return input.split(';').map(s => toKey(s.trim())).filter(Boolean);
  }

  if (typeof input === 'object') {
    const o: any = input;
    return [toKey(o?.Id ?? o?.id ?? o?.Key ?? o?.value ?? o)];
  }

  return [toKey(input)];
}

function clampToExisting(values: string[], opts: OptionItem[]): string[] {
  const allowed = new Set(opts.map(o => toKey(o.key)));
  return values.filter(v => allowed.has(v));
}

function useOptionMaps(options: OptionItem[]) {
  return React.useMemo(() => {
    const keyToText = new Map<string, string>();
    const keyToNumber = new Map<string, number>();
    for (const o of options) {
      const keyStr = toKey(o.key);
      keyToText.set(keyStr, o.text);
      const maybeNum =
        typeof o.key === 'number'
          ? o.key
          : Number.isFinite(Number(keyStr))
          ? Number(keyStr)
          : NaN;
      if (!Number.isNaN(maybeNum)) keyToNumber.set(keyStr, maybeNum);
    }
    return { keyToText, keyToNumber };
  }, [options]);
}

// ---------- component ----------
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
    fieldType,
    className,
    description,
    submitting,
  } = props;

  const isMulti = !!(multiselect ?? multiSelect);
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
  } = React.useContext(DynamicFormContext as any);

  const [isRequired, setIsRequired] = React.useState<boolean>(!!requiredProp);
  const [isDisabled, setIsDisabled] = React.useState<boolean>(!!disabledProp);
  const [isHidden, setIsHidden] = React.useState<boolean>(false);

  const [selectedKeys, setSelectedKeys] = React.useState<string[]>([]);
  const [selectedKey, setSelectedKey] = React.useState<string | null>(null);
  const [error, setError] = React.useState<string>('');
  const [touched, setTouched] = React.useState<boolean>(false);

  const { keyToText, keyToNumber } = useOptionMaps(options);

  // Mirror UI error -> global error (null when empty) under correct internal name
  const reportError = React.useCallback(
    (msg: string) => {
      const targetId = isLookup ? `${id}Id` : id;
      setError(msg || '');
      GlobalErrorHandle(targetId, msg || null);
    },
    [GlobalErrorHandle, id, isLookup]
  );

  // reflect required/disabled props
  React.useEffect(() => {
    setIsRequired(!!requiredProp);
    setIsDisabled(!!disabledProp);
  }, [requiredProp, disabledProp]);

  // submitting disables (standalone effect, like your reference)
  React.useEffect(() => {
    if (submitting === true) setIsDisabled(true);
  }, [submitting]);

  // prefill + hide/disable via rules
  React.useEffect(() => {
    const ensureInOptions = (vals: string[]) => clampToExisting(vals, options);

    // Prefill: New vs Edit/View
    if (FormMode == 8) {
      if (isMulti) {
        const initArr = ensureInOptions(
          starterValue != null
            ? (Array.isArray(starterValue)
                ? starterValue.map(toKey)
                : [toKey(starterValue)])
            : []
        );
        setSelectedKeys(initArr);
        setSelectedKey(null);
      } else {
        const init = starterValue != null ? toKey(starterValue) : '';
        const clamped = ensureInOptions(init ? [init] : []);
        setSelectedKey(clamped[0] ?? null);
        setSelectedKeys([]);
      }
    } else {
      const raw = FormData
        ? (isLookup ? (FormData as any)[`${id}Id`] : (FormData as any)[id])
        : undefined;

      if (isMulti) {
        const arr = ensureInOptions(normalizeToStringArray(raw));
        setSelectedKeys(arr);
        setSelectedKey(null);
      } else {
        const arr = ensureInOptions(normalizeToStringArray(raw));
        setSelectedKey(arr[0] ?? null);
        setSelectedKeys([]);
      }
    }

    // Disable / Hide logic
    if (FormMode === 4) {
      // Display form: field is disabled
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

    // Clear errors on prefill
    reportError('');
    setTouched(false);
    // NOTE: GlobalErrorHandle intentionally not in deps
  }, [
    FormData,
    FormMode,
    starterValue,
    options,
    isLookup,
    id,
    isMulti,
    displayName,
    AllDisableFields,
    AllHiddenFields,
    userBasedPerms,
    curUserInfo,
    listCols,
    reportError,
  ]);

  // validate & commit
  const validate = React.useCallback((): string => {
    if (isRequired) {
      if (isMulti && selectedKeys.length === 0) return REQUIRED_MSG;
      if (!isMulti && !selectedKey) return REQUIRED_MSG;
    }
    return '';
  }, [isRequired, isMulti, selectedKeys, selectedKey]);

  const commitValue = React.useCallback(() => {
    const err = validate();
    reportError(err);

    const targetId = isLookup ? `${id}Id` : id;

    if (isLookup) {
      const valueForCommit = isMulti
        ? selectedKeys
            .map(k => keyToNumber.get(k))
            .filter((n): n is number => typeof n === 'number')
        : selectedKey
        ? keyToNumber.get(selectedKey) ?? null
        : null;
      GlobalFormData(targetId, valueForCommit);
    } else {
      const valueForCommit = isMulti ? selectedKeys : selectedKey ? selectedKey : null;
      GlobalFormData(targetId, valueForCommit);
    }
  }, [
    validate,
    reportError,
    GlobalFormData,
    id,
    isMulti,
    isLookup,
    selectedKeys,
    selectedKey,
    keyToNumber,
  ]);

  // selection handlers
  const handleOptionSelect = (
    _e: unknown,
    data: { optionValue?: string | number; selectedOptions: (string | number)[] }
  ) => {
    if (isMulti) {
      const next = (data.selectedOptions ?? []).map(toKey);
      setSelectedKeys(next);
      if (touched) reportError(isRequired && next.length === 0 ? REQUIRED_MSG : '');
    } else {
      const nextVal = data.optionValue != null ? toKey(data.optionValue) : null;
      setSelectedKey(nextVal);
      if (touched) reportError(isRequired && !nextVal ? REQUIRED_MSG : '');
    }
  };

  const handleBlur = () => {
    setTouched(true);
    commitValue();
  };

  // display text & disabled
  const selectedOptions = isMulti ? selectedKeys : selectedKey ? [selectedKey] : [];
  const displayText = isMulti
    ? selectedKeys.length
      ? selectedKeys.map(k => keyToText.get(k) ?? k).join('; ')
      : ''
    : selectedKey
    ? keyToText.get(selectedKey) ?? selectedKey
    : '';
  const effectivePlaceholder = displayText || placeholder;

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
          multiselect={isMulti}
          disabled={isDisabled}
          inlinePopup
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

        {description !== '' && (
          <div className="descriptionText">{description}</div>
        )}
      </Field>
    </div>
  );
}


