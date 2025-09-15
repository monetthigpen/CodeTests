import * as React from 'react';
import { Field, Dropdown, Option } from '@fluentui/react-components';
import { DynamicFormContext } from './DynamicFormContext';
import formFieldsSetup, { FormFieldsProps } from './formFieldBased';

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
  fieldType?: string;      // 'lookup' → `${id}Id` or `${id}LookupId` (auto)
  multiselect?: boolean;   // v9 prop
  multiSelect?: boolean;   // v8 alias
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

const clampToExisting = (values: string[], opts: Array<{ key: string | number }>): string[] => {
  const allowed = new Set(opts.map(o => toKey(o.key)));
  return values.filter(v => allowed.has(v));
};

// Prefer `${id}LookupId` if it exists in listCols; else `${id}Id`
const resolveLookupKey = (id: string, listCols: unknown): string => {
  const bag = (listCols ?? {}) as Record<string, unknown>;
  const k1 = `${id}LookupId`;
  const k2 = `${id}Id`;
  if (Object.prototype.hasOwnProperty.call(bag, k1)) return k1;
  if (Object.prototype.hasOwnProperty.call(bag, k2)) return k2;
  return k2;
};

const buildCommitValue = (isLookup: boolean, isMulti: boolean, keys: string[]): unknown => {
  if (keys.length === 0) return null;

  if (!isLookup) return isMulti ? keys : keys[0];

  const nums = keys.map(k => Number(k)).filter((n): n is number => Number.isFinite(n));
  if (nums.length === 0) return null;

  return isMulti ? { results: nums } : nums[0];
};

export default function DropdownComponent(props: DropdownProps): JSX.Element {
  const {
    id, displayName, options, starterValue,
    isRequired = false, placeholder, className, description,
    fieldType, multiselect, multiSelect,
    disabled = false, submitting = false,
  } = props;

  const isLookup = fieldType === 'lookup';
  const isMulti = !!(multiselect ?? multiSelect);

  const [localVal, setLocalVal] = React.useState<string>('');
  const [selectedKeys, setSelectedKeys] = React.useState<string[]>([]);
  const [error, setError] = React.useState<string>('');
  const [isDisabled, setIsDisabled] = React.useState<boolean>(!!disabled);
  const [isHidden, setIsHidden] = React.useState<boolean>(false);

  const {
    FormData, GlobalFormData, GlobalErrorHandle, FormMode,
    AllDisableFields, AllHiddenFields, userBasedPerms, curUserInfo, listCols,
  } = React.useContext(DynamicFormContext);

  const targetId = React.useMemo(
    () => (isLookup ? resolveLookupKey(id, listCols) : id),
    [isLookup, id, listCols]
  );

  const keyToText = React.useMemo(() => {
    const m = new Map<string, string>();
    for (const o of options) m.set(toKey(o.key), o.text);
    return m;
  }, [options]);

  const textFromKeys = React.useCallback(
    (arr: string[]) => arr.map(k => keyToText.get(k) ?? k).join(';'),
    [keyToText]
  );

  // Init + prefill + rules
  React.useEffect(() => {
    let initKeys: string[] = [];
    if (FormMode === 8) {
      initKeys =
        starterValue == null
          ? []
          : Array.isArray(starterValue)
          ? (starterValue as (string | number)[]).map(toKey)
          : [toKey(starterValue)];
    } else {
      const bag = (FormData ?? {}) as Record<string, unknown>;
      const raw = bag[targetId];
      initKeys = normalizeToStringArray(raw);
    }
    initKeys = clampToExisting(initKeys, options);
    setSelectedKeys(initKeys);
    setLocalVal(textFromKeys(initKeys));

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

    setError('');
    GlobalErrorHandle(targetId, null);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []); // once

  // Disable on submit, keep visible text
  React.useEffect(() => {
    if (submitting) {
      setIsDisabled(true);
      setLocalVal(textFromKeys(selectedKeys));
    }
  }, [submitting, selectedKeys, textFromKeys]);

  // --- commit helper with explicit logging (remove logs after verifying) ---
  const commitNow = (keys: string[]): void => {
    const payload = buildCommitValue(isLookup, isMulti, keys);
    const hasError = isRequired && keys.length === 0;
    const errMsg = hasError ? REQUIRED_MSG : '';

    setError(errMsg);
    GlobalErrorHandle(targetId, hasError ? REQUIRED_MSG : null);
    GlobalFormData(targetId, payload);

    // DEBUG: verify the exact key & payload hitting context
    // eslint-disable-next-line no-console
    console.log('[Dropdown commit]', { targetId, isLookup, isMulti, payload });
  };

  const onOptionSelect = (_e: OnOptionSelectEvent, data: OnOptionSelectData): void => {
    const next = (data.selectedOptions ?? []).map(v => String(v));
    setSelectedKeys(next);
    setLocalVal(textFromKeys(next));
    commitNow(next); // commit immediately
  };

  const handleBlur = (): void => {
    commitNow(selectedKeys);
  };

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
            <Option key={String(o.key)} value={String(o.key)}>
              {o.text}
            </Option>
          ))}
        </Dropdown>

        {description ? <div className="descriptionText">{description}</div> : null}
      </Field>
    </div>
  );
}





