import * as React from 'react';
import { Field } from '@fluentui/react-components';         // v9 Field for label/validation
import { TagPicker, ITag } from '@fluentui/react';          // v8 TagPicker (chips)
import { DynamicFormContext } from './DynamicFormContext';

// ----- Hard-coded people (numeric IDs) -----
const CATALOG: Array<{ id: number; name: string; email?: string }> = [
  { id: 101, name: 'Ada Lovelace',     email: 'ada@example.com' },
  { id: 102, name: 'Alan Turing',      email: 'alan@example.com' },
  { id: 103, name: 'Grace Hopper',     email: 'grace@example.com' },
  { id: 104, name: 'Katherine Johnson',email: 'katherine@example.com' },
  { id: 105, name: 'Donald Knuth',     email: 'donald@example.com' },
];

export interface PeoplePickerProps {
  id: string;
  displayName: string;
  fieldType?: 'lookup' | string; // to match combobox commit behavior
  single?: boolean;

  starterValue?:
    | number
    | string
    | Array<number | string>
    | { results?: Array<number | string> }
    | null
    | undefined;

  isRequired?: boolean;
  disabled?: boolean;
  submitting?: boolean;
  placeholder?: string;
  className?: string;
  description?: string;
}

const REQUIRED_MSG = 'This is a required field and cannot be blank!';

const toKey = (n: number) => String(n);

const toNum = (v: unknown): number | null => {
  const n = typeof v === 'string' ? Number(v) : (v as number);
  return Number.isFinite(n) ? Number(n) : null;
};

function normalizeToIds(input: unknown): number[] {
  if (input == null) return [];

  if (
    typeof input === 'object' &&
    input !== null &&
    Array.isArray((input as { results?: Array<number | string> }).results)
  ) {
    return ((input as { results: Array<number | string> }).results)
      .map(toNum)
      .filter((n): n is number => n !== null);
  }

  if (Array.isArray(input)) {
    return (input as unknown[])
      .map(toNum)
      .filter((n): n is number => n !== null);
  }

  if (typeof input === 'string') {
    const parts = input.split(/[;,]/).map(s => s.trim()).filter(Boolean);
    return parts.map(toNum).filter((n): n is number => n !== null);
  }

  const n = toNum(input);
  return n === null ? [] : [n];
}

const arraysEqualTags = (a: ITag[], b: ITag[]) =>
  a.length === b.length && a.every((v, i) => v.key === b[i].key && v.name === b[i].name);

export default function PeoplePicker(props: PeoplePickerProps): JSX.Element {
  const {
    id,
    displayName,
    fieldType,
    single,
    starterValue,
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

  const [tags, setTags] = React.useState<ITag[]>([]);
  const didInitRef = React.useRef<boolean>(false);

  // Pass undefined instead of null to GlobalErrorHandle (fixes TS2345)
  const reportError = React.useCallback(
    (msg: string): void => {
      const out = msg || '';
      if (out !== error) setError(out);
      GlobalErrorHandle?.(id, out || undefined);
    },
    [GlobalErrorHandle, id, error]
  );

  React.useEffect(() => {
    if (isRequired !== !!requiredProp) setIsRequired(!!requiredProp);
    if (isDisabled !== !!disabledProp) setIsDisabled(!!disabledProp);
  }, [requiredProp, disabledProp, isRequired, isDisabled]);

  React.useEffect(() => {
    if (submitting === true && !isDisabled) setIsDisabled(true);
  }, [submitting, isDisabled]);

  // seed from starterValue once
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

  const onChange = (items?: ITag[]): void => {
    let next = items ?? [];
    if (single && next.length > 1) next = [next[next.length - 1]];
    setTags(next);
    if (touched) reportError(isRequired && next.length === 0 ? REQUIRED_MSG : '');
  };

  const validate = React.useCallback((): string => {
    if (isRequired && tags.length === 0) return REQUIRED_MSG;
    return '';
  }, [isRequired, tags]);

  const commitValue = React.useCallback((): void => {
    const err = validate();
    reportError(err);

    const targetId = isLookup ? `${id}Id` : id;
    const ids = tags.map(t => Number(t.key)).filter(n => Number.isFinite(n));

    if (single) {
      (GlobalFormData as (name: string, value: unknown) => void)?.(targetId, ids.length ? ids[0] : undefined);
    } else {
      (GlobalFormData as (name: string, value: unknown) => void)?.(targetId, ids);
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


import PeoplePicker from './PeoplePicker';


case "user": {
  allFormElements.push(
    <PeoplePicker
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
