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

type OnSelect = NonNullable<React.ComponentProps<typeof Dropdown>['onOptionSelect']>;
type OnSelectEvent = Parameters<OnSelect>[0];
type OnSelectData = Parameters<OnSelect>[1];

interface RuleResult {
  isDisabled?: boolean;
  isHidden?: boolean;
}

const REQUIRED_MSG = 'This is a required field and cannot be blank!';
const toKey = (k: unknown): string => (k == null ? '' : String(k));

const normalizeValues = (input: unknown): string[] => {
  if (!input) return [];
  if (Array.isArray(input)) return input.map(toKey);
  if (typeof input === 'object' && input !== null && (input as { results?: unknown[] }).results) {
    return ((input as { results: unknown[] }).results!).map(toKey);
  }
  if (typeof input === 'string' && input.includes(';')) {
    return input.split(';').map(s => s.trim()).filter(Boolean);
  }
  return [toKey(input)];
};

const resolveLookupKey = (id: string, listCols: unknown): string => {
  const cols = (listCols ?? {}) as Record<string, unknown>;
  return Object.prototype.hasOwnProperty.call(cols, `${id}LookupId`) ? `${id}LookupId` : `${id}Id`;
};

const buildValue = (isLookup: boolean, isMulti: boolean, keys: string[]): unknown => {
  if (!keys.length) return null;
  if (!isLookup) return isMulti ? keys : keys[0];
  const nums = keys.map(Number).filter(n => Number.isFinite(n));
  return nums.length ? (isMulti ? { results: nums } : nums[0]) : null;
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
    const map = new Map<string, string>(options.map(o => [toKey(o.key), o.text]));
    return (arr: string[]) => arr.map(k => map.get(k) ?? k).join(';');
  }, [options]);

  // Prefill + rules
  React.useEffect(() => {
    const bag = (FormData ?? {}) as Record<string, unknown>;
    let init = FormMode === 8 ? normalizeValues(starterValue) : normalizeValues(bag[targetId]);
    const allowed = new Set(options.map(o => toKey(o.key)));
    init = init.filter(v => allowed.has(v));

    setSelectedKeys(init);
    setLocalVal(keyToText(init));

    if (FormMode === 4) {
      setIsDisabled(true);
    } else {
      const ruleArgs: FormFieldsProps = {
        disabledList: AllDisableFields,
        hiddenList: AllHiddenFields,
        userBasedList: userBasedPerms,
        curUserList: curUserInfo,
        curField: displayName,
        formStateData: FormData,
        listColumns: listCols,
      } as unknown as FormFieldsProps;

      const results: RuleResult[] = (formFieldsSetup(ruleArgs) as RuleResult[]) || [];
      for (const r of results) {
        if (r.isDisabled !== undefined) setIsDisabled(!!r.isDisabled);
        if (r.isHidden !== undefined) setIsHidden(!!r.isHidden);
      }
    }

    setError('');
    GlobalErrorHandle(targetId, null);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  // Disable on submit & keep joined text
  React.useEffect(() => {
    if (submitting) {
      setIsDisabled(true);
      setLocalVal(keyToText(selectedKeys));
    }
  }, [submitting, selectedKeys, keyToText]);

  const commit = (keys: string[]): void => {
    const val = buildValue(isLookup, isMulti, keys);
    const msg = isRequired && !keys.length ? REQUIRED_MSG : '';
    setError(msg);
    GlobalErrorHandle(targetId, msg || null);
    GlobalFormData(targetId, val);
  };

  const onSelect = (_e: OnSelectEvent, data: OnSelectData): void => {
    const next = (data.selectedOptions ?? []).map(String);
    setSelectedKeys(next);
    setLocalVal(keyToText(next));
    commit(next);
  };

  const rootClass = className ?? 'fieldClass';

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
            className={rootClass}
            title={localVal}
          />
        )}

        <Dropdown
          id={id}
          className={rootClass}
          multiselect={isMulti}
          inlinePopup
          disabled={isDisabled}
          value={localVal}
          placeholder={localVal || placeholder}
          selectedOptions={selectedKeys}
          onOptionSelect={onSelect}
          onBlur={() => commit(selectedKeys)}
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







