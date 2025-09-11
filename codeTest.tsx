import * as React from 'react';
import { Field, Dropdown, Option } from '@fluentui/react-components';
import { DynamicFormContext } from './DynamicFormContext';
import formFieldsSetup, { FormFieldsProps } from './formFieldBased';

type OptionItem = { key: string | number; text: string };

interface DropdownProps {
  id: string;
  starterValue: string;          // matching your TextArea style
  displayName: string;
  isRequired: boolean;
  placeholder: string;
  multiSelect: boolean;
  fieldType: string;             // "lookup" => commit under `${id}Id`
  options: OptionItem[];
  className: string;
  description: string;
  disabled: boolean;
  submitting: boolean;
}

const REQUIRED_MSG = 'This is a required field and cannot be blank!';

// -- helpers --
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

export default function DropdownComponent(props: DropdownProps): JSX.Element {
  const {
    id,
    starterValue,
    displayName,
    isRequired: requiredProp,
    placeholder,
    multiSelect,
    fieldType,
    options,
    className,
    description,
    disabled: disabledProp,
    submitting,
  } = props;

  const isMulti = !!multiSelect;
  const isLookup = fieldType === 'lookup';

  // 🔹 Context set up EXACTLY like your TextArea screenshot
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

  const [selectedKeys, setSelectedKeys] = React.useState<string[]>([]);
  const [selectedKey, setSelectedKey] = React.useState<string | null>(null);

  const [error, setError] = React.useState<string>('');
  const [touched, setTouched] = React.useState<boolean>(false);

  const { keyToText, keyToNumber } = useOptionMaps(options);

  // mirror UI error -> global (null when empty) under the correct internal name
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

  // submitting => disable (standalone effect like your TextArea)
  React.useEffect(() => {
    if (submitting === true) setIsDisabled(true);
  }, [submitting]);

  // Prefill + Disable/Hide rules
  React.useEffect(() => {
    const ensureInOptions = (vals: string[]) => clampToExisting(vals, options);

    // Prefill: New vs Edit/View
    if (FormMode == 8) {
      if (isMulti) {
        const initArr = ensureInOptions(starterValue ? [toKey(starterValue)] : []);
        setSelectedKeys(initArr);
        setSelectedKey(null);
      } else {
        const init = starterValue ? toKey(starterValue) : '';
        const clamped = ensureInOptions(init ? [init] : []);
        setSelectedKey(clamped[0] ?? null);
        setSelectedKeys([]);
      }
    } else {
      const raw = FormData
        ? (isLookup ? (FormData as any)[`${id}Id`] : (FormData as any)[id]) // eslint-disable-line @typescript-eslint/no-explicit-any
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

    // Disable or enable the field
    // For Display form, field is disabled
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
      } as any; // eslint-disable-line @typescript-eslint/no-explicit-any

      const results = formFieldsSetup(formFieldProps) || [];
      if (results.length > 0) {
        for (let i = 0; i < results.length; i++) {
          if (results[i].isDisabled !== undefined) setIsDisabled(results[i].isDisabled);
          if (results[i].isHidden !== undefined) setIsHidden(results[i].isHidden);
        }
      }
    }

    // clear errors on prefill
    reportError('');
    setTouched(false);
    // GlobalErrorHandle intentionally not in deps
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
  }, [validate, reportError, GlobalFormData, id, isMulti, isLookup, selectedKeys, selectedKey, keyToNumber]);

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

  // UI text & rendering
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


