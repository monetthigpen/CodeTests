import * as React from 'react';
import { Field, Dropdown, Option, Input } from '@fluentui/react-components';
import { DynamicFormContext } from './DynamicFormContext';
import formFieldsSetup, { FormFieldsProps } from './formFieldBased';

export interface DropdownProps {
  id: string;
  displayName: string;
  options: { key: string | number; text: string }[];
  starterValue?: string | number | Array<string | number>;
  isRequired?: boolean;
  placeholder?: string;
  className?: string;
  description?: string;
  fieldType?: string;      // "lookup"
  multiselect?: boolean;   // v9
  multiSelect?: boolean;   // v8 alias
  disabled?: boolean;
  submitting?: boolean;
}

const REQUIRED_MSG = 'This is a required field and cannot be blank!';
const toKey = (k: unknown): string => (k == null ? '' : String(k));

const normalizeValues = (input: unknown): string[] => {
  if (!input) return [];
  if (Array.isArray(input)) return input.map(toKey);
  if (typeof input === 'object' && (input as any).results) return (input as any).results.map(toKey);
  if (typeof input === 'string' && input.includes(';')) return input.split(';').map(s => s.trim());
  return [toKey(input)];
};

const resolveLookupKey = (id: string, listCols: any): string => {
  if (listCols?.[`${id}LookupId`] !== undefined) return `${id}LookupId`;
  return `${id}Id`;
};

const buildValue = (isLookup: boolean, isMulti: boolean, keys: string[]): any => {
  if (!keys.length) return null;
  if (!isLookup) return isMulti ? keys : keys[0];
  const nums = keys.map(Number).filter(n => !isNaN(n));
  return isMulti ? { results: nums } : nums[0];
};

export default function DropdownComponent({
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
}: DropdownProps): JSX.Element {
  const isLookup = fieldType === 'lookup';
  const isMulti = !!(multiselect ?? multiSelect);

  const [selectedKeys, setSelectedKeys] = React.useState<string[]>([]);
  const [localVal, setLocalVal] = React.useState('');
  const [error, setError] = React.useState('');
  const [isDisabled, setIsDisabled] = React.useState(disabled);
  const [isHidden, setIsHidden] = React.useState(false);

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

  const targetId = React.useMemo(
    () => (isLookup ? resolveLookupKey(id, listCols) : id),
    [isLookup, id, listCols]
  );

  const keyToText = React.useMemo(() => {
    const m = new Map(options.map(o => [toKey(o.key), o.text]));
    return (arr: string[]) => arr.map(k => m.get(k) ?? k).join(';');
  }, [options]);

  // Prefill + rules
  React.useEffect(() => {
    let init = FormMode === 8 ? normalizeValues(starterValue) : normalizeValues((FormData ?? {})[targetId]);
    init = init.filter(v => options.some(o => toKey(o.key) === v));

    setSelectedKeys(init);
    setLocalVal(keyToText(init));

    if (FormMode === 4) setIsDisabled(true);
    else {
      const rules: FormFieldsProps = {
        disabledList: AllDisableFields,
        hiddenList: AllHiddenFields,
        userBasedList: userBasedPerms,
        curUserList: curUserInfo,
        curField: displayName,
        formStateData: FormData,
        listColumns: listCols,
      } as any;
      (formFieldsSetup(rules) || []).forEach(r => {
        if (r.isDisabled !== undefined) setIsDisabled(!!r.isDisabled);
        if (r.isHidden !== undefined) setIsHidden(!!r.isHidden);
      });
    }

    setError('');
    GlobalErrorHandle(targetId, null);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  // Disable on submit
  React.useEffect(() => {
    if (submitting) {
      setIsDisabled(true);
      setLocalVal(keyToText(selectedKeys));
    }
  }, [submitting, selectedKeys, keyToText]);

  const commit = (keys: string[]) => {
    const val = buildValue(isLookup, isMulti, keys);
    const msg = isRequired && !keys.length ? REQUIRED_MSG : '';
    setError(msg);
    GlobalErrorHandle(targetId, msg || null);
    GlobalFormData(targetId, val);
  };

  const onSelect = (_: any, data: any) => {
    const next = (data.selectedOptions ?? []).map(String);
    setSelectedKeys(next);
    setLocalVal(keyToText(next));
    commit(next);
  };

  return (
    <div style={{ display: isHidden ? 'none' : 'block' }}>
      <Field
        label={displayName}
        required={isRequired}
        validationMessage={error || undefined}
        validationState={error ? 'error' : undefined}
      >
        {isDisabled && (
          <Input
            id={id}
            disabled
            value={localVal}
            placeholder={localVal || placeholder}
            className={className ?? 'fieldClass'}
            title={localVal}
          />
        )}
        <Dropdown
          id={id}
          className={className ?? 'fieldClass'}
          multiselect={isMulti}
          inlinePopup
          disabled={isDisabled}
          value={localVal}
          placeholder={localVal || placeholder}
          selectedOptions={selectedKeys}
          onOptionSelect={onSelect}
          onBlur={() => commit(selectedKeys)}
          title={localVal}
        >
          {options.map(o => (
            <Option key={String(o.key)} value={String(o.key)}>
              {o.text}
            </Option>
          ))}
        </Dropdown>
        {description && <div className="descriptionText">{description}</div>}
      </Field>
    </div>
  );
}







