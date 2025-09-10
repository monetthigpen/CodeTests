import * as React from 'react';
import { Field, Combobox, Option } from '@fluentui/react-components';
import { DynamicFormContext } from './DynamicFormContext';

type PersonOption = {
  key: string | number;   // person/user id (string or number)
  text: string;           // primary display (e.g., Full Name)
  subText?: string;       // optional (e.g., email)
};

export interface PeoplePickerProps {
  id: string;
  displayName: string;
  placeholder?: string;
  isRequired?: boolean;
  disabled?: boolean;
  /** If true => multi; otherwise single */
  multiselect?: boolean;
  /** Initial options (e.g., first page or known users) */
  options?: PersonOption[];
  /** Starter value(s) for "New" form */
  starterValue?: string | number | Array<string | number>;
  /** When provided, used to fetch suggestions as user types */
  onSearch?: (query: string) => Promise<PersonOption[]>;
  /** SP/SharePoint person fields should commit numeric IDs */
  fieldType?: 'user' | 'lookup' | string;
  className?: string;
  description?: string;
  submitting?: boolean; // submitting disables via its own effect
}

// ---------------- helpers ----------------

const toKey = (k: unknown): string => (k == null ? '' : String(k));

function normalizeToStringArray(input: unknown): string[] {
  if (input == null) return [];

  if (Array.isArray((input as any)?.results)) {
    return ((input as any).results as unknown[]).map(toKey);
  }

  if (Array.isArray(input)) {
    const arr = input as unknown[];
    if (arr.length && typeof arr[0] === 'object' && arr[0] !== null) {
      return arr.map((o: any) => toKey(o?.Id ?? o?.id ?? o?.Key ?? o?.value ?? o)); // eslint-disable-line
    }
    return arr.map(toKey);
  }

  if (typeof input === 'string' && input.includes(';')) {
    return input.split(';').map(s => toKey(s.trim())).filter(Boolean);
  }

  if (typeof input === 'object') {
    const o: any = input; // eslint-disable-line
    return [toKey(o?.Id ?? o?.id ?? o?.Key ?? o?.value ?? o)];
  }

  return [toKey(input)];
}

function clampToExisting(values: string[], opts: PersonOption[]): string[] {
  const allowed = new Set(opts.map(o => toKey(o.key)));
  return values.filter(v => allowed.has(v));
}

function useOptionMaps(options: PersonOption[]) {
  return React.useMemo(() => {
    const keyToText = new Map<string, string>();
    const keyToSub = new Map<string, string | undefined>();
    const keyToNumber = new Map<string, number>();

    for (const o of options) {
      const k = toKey(o.key);
      keyToText.set(k, o.text);
      keyToSub.set(k, o.subText);
      const maybeNum = typeof o.key === 'number' ? o.key : Number.isFinite(Number(k)) ? Number(k) : NaN;
      if (!Number.isNaN(maybeNum)) keyToNumber.set(k, maybeNum);
    }
    return { keyToText, keyToSub, keyToNumber };
  }, [options]);
}

// ---------------- component ----------------

export default function PeoplePickerComponent(props: PeoplePickerProps): JSX.Element {
  const {
    id,
    displayName,
    placeholder,
    isRequired: requiredProp,
    disabled: disabledProp,
    multiselect,
    options = [],
    starterValue,
    onSearch,
    fieldType,
    className,
    description,
    submitting,
  } = props;

  const isMulti = !!multiselect;
  const isUserLike = props.fieldType === 'user' || props.fieldType === 'lookup';

  const { FormData, GlobalFormData, FormMode, GlobalErrorHandle } =
    React.useContext(DynamicFormContext);

  const [isRequired, setIsRequired] = React.useState<boolean>(!!requiredProp);
  const [isDisabled, setIsDisabled] = React.useState<boolean>(!!disabledProp);
  const [error, setError] = React.useState<string>('');
  const [touched, setTouched] = React.useState<boolean>(false);

  // Local options: initial + async search results
  const [localOptions, setLocalOptions] = React.useState<PersonOption[]>(options);

  // selections are stored as string keys to match Option values
  const [selectedKeys, setSelectedKeys] = React.useState<string[]>([]);      // multi
  const [selectedKey, setSelectedKey] = React.useState<string | null>(null); // single

  // maps based on current localOptions
  const { keyToText, keyToSub, keyToNumber } = useOptionMaps(localOptions);

  // single place to mirror UI error -> global error (null when empty)
  const reportError = React.useCallback(
    (msg: string) => {
      setError(msg || '');
      GlobalErrorHandle(id, msg || null);
    },
    [GlobalErrorHandle, id]
  );

  // reflect external required/disabled props
  React.useEffect(() => {
    setIsRequired(!!requiredProp);
    setIsDisabled(!!disabledProp);
  }, [requiredProp, disabledProp]);

  // submitting disables (own effect)
  React.useEffect(() => {
    if (submitting === true) setIsDisabled(true);
  }, [submitting]);

  // keep localOptions in sync if props.options changes
  React.useEffect(() => {
    setLocalOptions(options);
  }, [options]);

  // prefill (New vs Edit/View)
  React.useEffect(() => {
    const ensureInOptions = (vals: string[]) => clampToExisting(vals, localOptions);

    if (FormMode == 8) {
      if (isMulti) {
        const initArr =
          starterValue != null
            ? (Array.isArray(starterValue) ? starterValue.map(toKey) : [toKey(starterValue)])
            : [];
        const clamped = ensureInOptions(initArr);
        setSelectedKeys(clamped);
        setSelectedKey(null);
      } else {
        const init = starterValue != null ? toKey(starterValue) : '';
        const clamped = ensureInOptions(init ? [init] : []);
        setSelectedKey(clamped[0] ?? null);
        setSelectedKeys([]);
      }
    } else {
      const raw = FormData
        ? (isUserLike ? (FormData as any)[`${id}Id`] : (FormData as any)[id])
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

    // clear any previous errors
    reportError('');
    setTouched(false);
  }, [FormData, FormMode, starterValue, localOptions, isUserLike, id, isMulti, reportError]);

  // ---- search handling (optional) ----
  const handleInputChange = React.useCallback(
    async (_e: unknown, data: { value: string }) => {
      const q = data?.value ?? '';
      if (!onSearch) return;                 // static options only
      try {
        const results = await onSearch(q);   // expected as PersonOption[]
        // merge by key (keep unique)
        const map = new Map<string, PersonOption>();
        for (const o of results) map.set(toKey(o.key), o);
        for (const o of localOptions) map.set(toKey(o.key), o);
        setLocalOptions(Array.from(map.values()));
      } catch {
        // ignore search errors for now
      }
    },
    [onSearch, localOptions]
  );

  // ---- validation & commit ----
  const validate = React.useCallback((): string => {
    if (isRequired) {
      if (isMulti && selectedKeys.length === 0) return 'This is a required field and cannot be blank!';
      if (!isMulti && !selectedKey) return 'This is a required field and cannot be blank!';
    }
    return '';
  }, [isRequired, isMulti, selectedKeys, selectedKey]);

  const commitValue = React.useCallback(() => {
    const err = validate();
    reportError(err);

    if (isUserLike) {
      // Commit numeric IDs (SharePoint person/lookup pattern)
      const valueForCommit = isMulti
        ? selectedKeys
            .map(k => keyToNumber.get(k))
            .filter((n): n is number => typeof n === 'number')
        : selectedKey
        ? keyToNumber.get(selectedKey) ?? null
        : null;
      GlobalFormData(id, valueForCommit);
    } else {
      // Non-user: commit string(s); single uses null when empty
      const valueForCommit = isMulti ? selectedKeys : (selectedKey ? selectedKey : null);
      GlobalFormData(id, valueForCommit);
    }
  }, [validate, reportError, isUserLike, isMulti, selectedKeys, selectedKey, keyToNumber, GlobalFormData, id]);

  // ---- selection handlers ----
  const handleOptionSelect = (
    _e: unknown,
    data: { selectedOptions: (string | number)[]; optionValue?: string | number }
  ) => {
    if (isMulti) {
      const next = (data.selectedOptions ?? []).map(toKey);
      setSelectedKeys(next);
      if (touched) reportError(isRequired && next.length === 0 ? 'This is a required field and cannot be blank!' : '');
    } else {
      const nextVal = data.optionValue != null ? toKey(data.optionValue) : null;
      setSelectedKey(nextVal);
      if (touched) reportError(isRequired && !nextVal ? 'This is a required field and cannot be blank!' : '');
    }
  };

  const handleBlur = () => {
    setTouched(true);
    commitValue();
  };

  // ---- display text (semicolon for multi) ----
  const selectedOptions = isMulti ? selectedKeys : (selectedKey ? [selectedKey] : []);
  const displayText = isMulti
    ? (selectedKeys.length ? selectedKeys.map(k => keyToText.get(k) ?? k).join('; ') : '')
    : (selectedKey ? (keyToText.get(selectedKey) ?? selectedKey) : '');
  const effectivePlaceholder = displayText || placeholder;

  const hasError = !!error;

  return (
    <Field
      label={displayName}
      required={isRequired}
      validationMessage={hasError ? error : undefined}
      validationState={hasError ? 'error' : undefined}
    >
      <Combobox
        id={id}
        placeholder={effectivePlaceholder}
        disabled={isDisabled}
        multiselect={isMulti}
        // typing/search
        onInputChange={handleInputChange}
        // selection
        selectedOptions={selectedOptions}
        onOptionSelect={handleOptionSelect}
        onBlur={handleBlur}
        className={className}
      >
        {localOptions.map(o => {
          const value = toKey(o.key);
          const label = o.text;
          const secondary = keyToSub.get(value);
        return (
            <Option key={value} value={value} text={label}>
              {secondary ? `${label} — ${secondary}` : label}
            </Option>
          );
        })}
      </Combobox>

      {description !== '' && <div className="descriptionText">{description}</div>}
    </Field>
  );
}
