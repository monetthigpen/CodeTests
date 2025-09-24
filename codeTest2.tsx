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

export interface TagPeoplePickerSimpleProps {
  id: string;
  displayName: string;

  /** If true, only one person can be selected */
  single?: boolean;

  /** Validation / UX */
  isRequired?: boolean;
  disabled?: boolean;
  submitting?: boolean;             // when true, disables via its own effect
  placeholder?: string;
  className?: string;
  description?: string;
}

const REQUIRED_MSG = 'This is a required field and cannot be blank!';

const toKey = (n: number) => String(n);

/**
 * Simple People Picker using TagPicker with hard-coded values.
 * Commits numeric IDs via GlobalFormData:
 *  - single: number | null
 *  - multi:  number[]
 */
export default function TagPeoplePickerSimple(props: TagPeoplePickerSimpleProps): JSX.Element {
  const {
    id,
    displayName,
    single,
    isRequired: requiredProp,
    disabled: disabledProp,
    submitting,
    placeholder,
    className,
    description,
  } = props;

  const { GlobalFormData, GlobalErrorHandle } = React.useContext(DynamicFormContext);

  const [isRequired, setIsRequired] = React.useState<boolean>(!!requiredProp);
  const [isDisabled, setIsDisabled] = React.useState<boolean>(!!disabledProp);
  const [error, setError] = React.useState<string>('');
  const [touched, setTouched] = React.useState<boolean>(false);

  // Selected tags (TagPicker uses ITag = { key: string | number; name: string })
  const [tags, setTags] = React.useState<ITag[]>([]);

  // Mirror UI error -> global error (null when empty)
  const reportError = React.useCallback(
    (msg: string) => {
      setError(msg || '');
      GlobalErrorHandle(id, msg || null);
    },
    [GlobalErrorHandle, id]
  );

  // Reflect external flags
  React.useEffect(() => {
    setIsRequired(!!requiredProp);
    setIsDisabled(!!disabledProp);
  }, [requiredProp, disabledProp]);

  // Submitting disables (own effect like your TextArea ref)
  React.useEffect(() => {
    if (submitting === true) setIsDisabled(true);
  }, [submitting]);

  // Suggestions from hard-coded catalog
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

  // Change handler
  const onChange = (items?: ITag[]) => {
    let next = items ?? [];
    if (single && next.length > 1) {
      // enforce single: keep only the last added
      next = [next[next.length - 1]];
    }
    setTags(next);
    if (touched) reportError(isRequired && next.length === 0 ? REQUIRED_MSG : '');
  };

  // Validate + commit numeric IDs
  const validate = React.useCallback((): string => {
    if (isRequired && tags.length === 0) return REQUIRED_MSG;
    return '';
  }, [isRequired, tags]);

  const commitValue = React.useCallback(() => {
    const err = validate();
    reportError(err);

    const ids = tags.map(t => Number(t.key)).filter(n => Number.isFinite(n));
    GlobalFormData(id, single ? (ids[0] ?? null) : ids);
  }, [validate, reportError, tags, GlobalFormData, id, single]);

  const handleBlur = () => {
    setTouched(true);
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

      {description !== '' && <div className="descriptionText">{description}</div>}
    </Field>
  );
}
