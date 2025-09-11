import * as React from 'react';
import { Field } from '@fluentui/react-components';
import { Dropdown, IDropdownOption } from '@fluentui/react';
import { DynamicFormContext } from './DynamicFormContext';
import formFieldsSetup, { FormFieldsProps } from './formFieldBased';

interface OptionItem extends IDropdownOption {
  // key: string | number; text: string; (already in IDropdownOption)
}

interface DropdownProps {
  id: string;
  starterValue?: string | number | Array<string | number>;
  displayName: string;
  isRequired?: boolean;
  placeholder?: string;
  multiSelect?: boolean;
  fieldType?: string;                  // 'lookup' to commit under `${id}Id` as numbers
  options: OptionItem[];
  className?: string;
  description?: string;
  disabled?: boolean;
  submitting?: boolean;
}

const REQUIRED_MSG = 'This is a required field and cannot be blank!';

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

  const isMulti = !!multiSelect;
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

  // v8 selection state (strings for internal handling)
  const [selectedKeys, setSelectedKeys] = React.useState<string[]>([]);
  const [selectedKey, setSelectedKey] = React.useState<string | null>(null);

  const [error, setError] = React.useState<string>('');
  const [touched, setTouched] = React.useState<boolean>(false);

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

  // Submitting → disable
  React.useEffect(() => {
    if (submitting === true) setIsDisabled(true);
  }, [submitting]);

  // Prefill + rules + display mode
  React.useEffect(() => {
    const ensureInOptions = (vals: string[]) => clampToExisting(vals, options);

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

    // Disable or enable the field
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
    isMulti,
    displayName,
    AllDisableFields,
    AllHiddenFields,
    userBasedPerms,
    curUserInfo,
    listCols,
    reportError,
  ]);

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
      const toNum = (k: string | null): number | null => {
        if (k == null) return null;
        const n = Number(k);
        return Number.isFinite(n) ? n : null;
      };

      const valueForCommit = isMulti
        ? selectedKeys
            .map(k => toNum(k))
            .filter((n): n is number => n !== null)
        : toNum(selectedKey);

      GlobalFormData(targetId, valueForCommit);
    } else {
      const valueForCommit = isMulti ? selectedKeys : (selectedKey ? selectedKey : null);
      GlobalFormData(targetId, valueForCommit);
    }
  }, [validate, reportError, GlobalFormData, id, isMulti, isLookup, selectedKeys, selectedKey]);

  // v8 onChange handler
  const handleChange = (
    _e: React.FormEvent<HTMLElement | HTMLDivElement>,
    option?: IDropdownOption
  ) => {
    if (!option) return;
    const k = toKey(option.key);

    if (isMulti) {
      const next = option.selected
        ? [...selectedKeys, k]
        : selectedKeys.filter(x => x !== k);
      setSelectedKeys(next);
      if (touched) reportError(isRequired && next.length === 0 ? REQUIRED_MSG : '');
    } else {
      setSelectedKey(k);
      if (touched) reportError(isRequired && !k ? REQUIRED_MSG : '');
    }
  };

  const handleBlur = () => {
    setTouched(true);
    commitValue();
  };

  // text shown as placeholder for your UX (v8 renders selection itself)
  const displayText = isMulti
    ? (selectedKeys.length ? selectedKeys.join('; ') : '')
    : (selectedKey ?? '');

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
          multiSelect={isMulti}
          options={options}
          disabled={isDisabled}
          selectedKeys={isMulti ? selectedKeys : undefined}
          selectedKey={!isMulti ? (selectedKey ?? undefined) : undefined}
          onChange={handleChange}
          onBlur={handleBlur}
          className={className}
        />
        {description !== '' && (
          <div className="descriptionText">{description}</div>
        )}
      </Field>
    </div>
  );
}


