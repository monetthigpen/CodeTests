import * as React from 'react';
import { Field } from '@fluentui/react-components';
import { Dropdown, IDropdownOption } from '@fluentui/react';
import { DynamicFormContext } from './DynamicFormContext';
import formFieldsSetup, { FormFieldsProps } from './formFieldBased';

/* ------------------------------ Props ------------------------------ */

interface DropdownProps {
  id: string;
  starterValue?: string | number | Array<string | number>;
  displayName: string;
  isRequired?: boolean;
  placeholder?: string;
  /** Accept either prop name from callers */
  multiSelect?: boolean;      // Fluent v8 canonical prop
  multiselect?: boolean;      // compatibility alias
  fieldType?: string;         // 'lookup' => commit under `${id}Id` as numbers
  options: IDropdownOption[]; // Fluent v8 options
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

  // SharePoint-style multi lookup: { results: [...] }
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

  // Delimited string fallback (e.g., "1;2;3")
  if (typeof input === 'string' && input.includes(';')) {
    return input.split(';').map(s => toKey(s.trim())).filter(Boolean);
  }

  if (typeof input === 'object') {
    const o: any = input;
    return [toKey(o?.Id ?? o?.id ?? o?.Key ?? o?.value ?? o)];
  }

  return [toKey(input)];
}

function clampToExisting(values: string[], opts: IDropdownOption[]): string[] {
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
    multiSelect,
    multiselect, // alias supported
    fieldType,
    options,
    className,
    description,
    disabled: disabledProp = false,
    submitting = false,
  } = props;

  // Accept either prop; prefer canonical multiSelect if provided
  const isMulti = !!(multiSelect ?? multiselect);
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

  // Fluent v8 selection state (keep as strings)
  const [selectedKeys, setSelectedKeys] = React.useState<string[]>([]);      // multi
  const [selectedKey, setSelectedKey] = React.useState<string | null>(null); // single

  const [error, setError] = React.useState<string>('');
  const [touched, setTouched] = React.useState<boolean>(false);

  // Mirror UI error to global (null when empty) under correct internal name
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

  // submitting => disable (own effect like your TextArea)
  React.useEffect(() => {
    if (submitting === true) setIsDisabled(true);
  }, [submitting]);

  // Prefill + rules + display mode
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

    // Clear errors on prefill
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

  // Validation + commit
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

  // v8 change/blur handlers
  const handleChange = (
    _e: React.FormEvent<HTMLElement | HTMLDivElement>,
    option?: IDropdownOption
  ) => {
    if (!option) return;
    const k = String(option.key);

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

  // For UX: semicolon-joined label (v8 also renders its own selection)
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
          multiSelect={isMulti}  // v8 prop; we accept multiselect as an alias via isMulti
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



