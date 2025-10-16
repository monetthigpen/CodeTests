import * as React from 'react';
import { Field } from '@fluentui/react-components';         // v9 Field for label/validation
import { TagPicker, ITag } from '@fluentui/react';          // v8 TagPicker (chips)
import { DynamicFormContext } from './DynamicFormContext';

// ----- Hard-coded people (numeric IDs) -----
const CATALOG: Array<{ id: number; name: string; email?: string }> = [
  { id: 101, name: 'Ada Lovelace',      email: 'ada@example.com' },
  { id: 102, name: 'Alan Turing',       email: 'alan@example.com' },
  { id: 103, name: 'Grace Hopper',      email: 'grace@example.com' },
  { id: 104, name: 'Katherine Johnson', email: 'katherine@example.com' },
  { id: 105, name: 'Donald Knuth',      email: 'donald@example.com' },
];

export interface TagPeoplePickerSimpleProps {
  id: string;
  displayName: string;

  /** Match combobox behavior: 'lookup' commits to `${id}Id` as numbers */
  fieldType?: 'lookup' | string;

  /** If true, only one person can be selected (combobox single-select analog) */
  single?: boolean;

  /** Pre-fill value(s) like your other components */
  starterValue?:
    | number
    | string
    | Array<number | string>
    | { results?: Array<number | string> }
    | null
    | undefined;

  /** Validation / UX */
  isRequired?: boolean;
  disabled?: boolean;
  submitting?: boolean;      // when true, disables (like your combobox)
  placeholder?: string;
  className?: string;
  description?: string;
}

const REQUIRED_MSG = 'This is a required field and cannot be blank!';

const toKey = (n: number) => String(n);

// ---------- helpers: mirror your combobox starter handling ----------
const toNum = (v: unknown): number | null => {
  const n = typeof v === 'string' ? Number(v) : (v as number);
  return Number.isFinite(n) ? Number(n) : null;
};

function normalizeToIds(input: unknown): number[] {
  if (input == null) return [];

  // REST multi: { results: [] }
  if (typeof input === 'object' && input !== null && Array.isArray((input as { results?: unknown[] }).results)) {
    return ((input as { results: unknown[] }).results)
      .map(toNum)
      .filter((n): n is number => n !== null);
  }

  // Array of values
  if (Array.isArray(input)) {
    return (input as unknown[])
      .map(toNum)
      .filter((n): n is number => n !== null);
  }

  // String list "1;2,3"
  if (typeof input === 'string') {
    const parts = input.split(/[;,]/).map(s => s.trim()).filter(Boolean);
    return parts.map(toNum).filter((n): n is number => n !== null);
  }

  // Scalar
  const n = toNum(input);
  return n === null ? [] : [n];
}

const arraysEqualTags = (a: ITag[], b: ITag[]) =>
  a.length === b.length && a.every((v, i) => v.key === b[i].key && v.name === b[i].name);

export default function TagPeoplePickerSimple(props: TagPeoplePickerSimpleProps): JSX.Element {
  const {
    id,
    displayName,
    fieldType,
    single,
    starterValue,                   // ✅ match combobox: prefill support
    isRequired: requiredProp,
    disabled: disabledProp,
    submitting,
    placeholder,
    className,
    description,
  } = props;

  const isLookup = fieldType === 'lookup';

  const { GlobalFormData, GlobalErrorHandle } = React.useContext(DynamicFormContext);

  const [isRequired, setIsRequired] = React.useState<boolean>(!!requiredProp);
  const [isDisabled, setIsDisabled] = React.useState<boolean>(!!disabledProp);
  const [error, setError] = React.useState<string>('');
  const [touched, setTouched] = React.useState<boolean>(false);

  // TagPicker selection (ITag = { key: string | number; name: string })
  const [tags, setTags] = React.useState<ITag[]>([]);
  const didInitRef = React.useRef<boolean>(false);

  // Error → global error (null clears)
  const reportError = React.useCallback(
    (msg: string): void => {
      const out = msg || '';
      if (out !== error) setError(out);
      GlobalErrorHandle?.(id, out || null);
    },
    [GlobalErrorHandle, id, error]
  );

  // Prop → state sync (like combobox)
  React.useEffect(() => {
    if (isRequired !== !!requiredProp) setIsRequired(!!requiredProp);
    if (isDisabled !== !!disabledProp) setIsDisabled(!!disabledProp);
  }, [requiredProp, disabledProp, isRequired, isDisabled]);

  // Submitting disables (same pattern as combobox)
  React.useEffect(() => {
    if (submitting === true && !isDisabled) setIsDisabled(true);
  }, [submitting, isDisabled]);

  // ✅ Seed from starterValue, like your combobox
  React.useEffect(() => {
    if (didInitRef.current) return;

    const ids = normalizeToIds(starterValue);
    if (ids.length === 0) {
      didInitRef.current = true;
      return;
    }

    const byId = new Map(CATALOG.map(p => [p.id, p]));
    const seeded: ITag[] = [];

    if (single) {
      const first = ids[0];
      const p = first != null ? byId.get(first) : undefined;
      if (p) seeded.push({ key: toKey(p.id), name: p.name });
    } else {
      for (const idNum of ids) {
        const p = byId.get(idNum);
        if (p) seeded.push({ key: toKey(p.id), name: p.name });
      }
    }

    if (seeded.length > 0) {
      setTags(prev => (arraysEqualTags(prev, seeded) ? prev : seeded));
    }
    didInitRef.current = true;
  }, [starterValue, single]);

  // Suggestions from hard-coded catalog (kept simple)
  const onResolveSuggestions = React.useCallback(
    (filterText: string, selectedItems?: ITag[]): ITag[] => {
      const taken = new Set((selectedItems ?? []).map(t => String(t.key)));
      const ft = filterText?.toLowerCase() ?? '';
      return CATALOG
        .filter(p =>
          !taken.has(toKey(p.id)) &&
          (!ft ||
            p.name.toLowerCase().includes(ft) ||
            p.email?.toLowerCase().includes(ft)))
        .map(p => ({ key: toKey(p.id), name: p.name }));
    },
    []
  );

  // Change handler (enforce single like combobox single-select)
  const onChange = (items?: ITag[]): void => {
    let next = items ?? [];
    if (single && next.length > 1) next = [next[next.length - 1]];
    setTags(next);
    if (touched) reportError(isRequired && next.length === 0 ? REQUIRED_MSG : '');
  };

  // Validate + commit (combobox-style commit on blur)
  const validate = React.useCallback((): string => {
    if (isRequired && tags.length === 0) return REQUIRED_MSG;
    return '';
  }, [isRequired, tags]);

  const commitValue = React.useCallback((): void => {
    const err = validate();
    reportError(err);

    // Match combobox: lookup fields commit to `${id}Id` as numbers; otherwise to `id`
    const targetId = isLookup ? `${id}Id` : id;
    const ids = tags.map(t => Number(t.key)).filter(n => Number.isFinite(n));

    if (single) {
      (GlobalFormData as (name: string, value: unknown) => void)(targetId, ids.length ? ids[0] : null);
    } else {
      (GlobalFormData as (name: string, value: unknown) => void)(targetId, ids);
    }
  }, [validate, reportError, tags, GlobalFormData, id, isLookup, single]);

  const handleBlur = (): void => {
    if (!touched) setTouched(true);
    commitValue();
  };

  const hasError = !!error;

  return (
    <Field
      label={displayName}
      required={isRequired}
      validationMessage={hasError ? error : undefined}
      validationState={hasError ? 'error' : undefined}
    >
      <div onBlur={handleBlur} className={className}>
        <TagPicker
          onResolveSuggestions={onResolveSuggestions}
          selectedItems={tags}
          onChange={onChange}
          inputProps={{ placeholder: placeholder || '' }}
          itemLimit={single ? 1 : undefined}
          disabled={isDisabled}
        />
      </div>

      {description ? <div className="descriptionText">{description}</div> : null}
    </Field>
  );
}


case "user": {
  allFormElements.push(
    <TagPeoplePickerSimple
      id={listColumns[i].name}
      displayName={listColumns[i].displayName}
      starterValue={starterVal}                  // initial value if provided
      isRequired={listColumns[i].required}
      submitting={isSubmitting}                  // disables when submitting
      single={!listColumns[i].multi}             // true if field is single-user
      placeholder={listColumns[i].description}   // optional
      description={listColumns[i].description}
      className="elementsWidth"
    />
  );
  break;
}
