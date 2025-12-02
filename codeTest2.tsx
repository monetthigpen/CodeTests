import * as React from "react";
import {
  TagPicker,
  TagPickerInput,
  makeStyles,
  shorthands,
  tokens,
  ITag
} from "@fluentui/react-components";

import { DynamicFormContext } from "../KS/DynamicFormKS";

// ---------- CONSTANTS ----------
const REQUIRED_MSG = "This field is required";

// ---------- STYLES ----------
const useStyles = makeStyles({
  root: {
    display: "flex",
    flexDirection: "column",
    ...shorthands.gap(tokens.spacingVerticalXXS),
  },
});

// ---------- PickerEntity Type ----------
export interface PickerEntity {
  Key: string;
  DisplayText: string;
  EntityData: {
    Email?: string;
    SPUserID: number;
  };
  EntityType: "User";
}

// ---------- Component Props ----------
export interface PeoplePickerProps {
  id: string;
  displayName: string;
  className?: string;
  description?: string;
  placeholder?: string;
  isRequired?: boolean;
  disabled?: boolean;
  multiselect?: boolean;
  submitting?: boolean;
  principalType?: number;
  maxSuggestions?: number;
  starterValue?: number | number[];

  // Fluent UI / SPFx dependencies
  spHttpClient: any;
  spHttpClientConfig: any;
}

// =========================================================
//                     PeoplePicker Component
// =========================================================
const PeoplePicker: React.FC<PeoplePickerProps> = (props) => {
  const {
    id,
    displayName,
    className,
    description,
    placeholder,
    isRequired,
    disabled,
    multiselect = false,
    principalType = 1,
    maxSuggestions = 5,
    starterValue,
    spHttpClient,
    spHttpClientConfig,
  } = props;

  const ctx = React.useContext(DynamicFormContext);
  const styles = useStyles();

  // ---------- Form Mode ----------
  const mode = ctx.FormMode; // 4=view, 6=edit, 8=new
  const isView = mode === 4;
  const isEdit = mode === 6;
  const isNew = mode === 8;

  // ---------- Component State ----------
  const [query, setQuery] = React.useState("");
  const [touched, setTouched] = React.useState(false);
  const [selectedOptions, setSelectedOptions] = React.useState<string[]>([]);
  const [options, setOptions] = React.useState<Map<string, PickerEntity>>(new Map());

  // ---------- Global Refs ----------
  const errRef = ctx.setGlobalError(id);
  const valueRef = ctx.setGlobalRef(id);

  // ---------- Commit Value to GlobalFormData ----------
  const commitValue = React.useCallback(() => {
    const entities: PickerEntity[] = selectedOptions
      .map((key) => options.get(key))
      .filter((v): v is PickerEntity => !!v);

    // Required validation
    if (isRequired && entities.length === 0) {
      errRef(REQUIRED_MSG);
    } else {
      errRef("");
    }

    valueRef(entities);
  }, [selectedOptions, options, isRequired, errRef, valueRef]);

  // =========================================================
  //                  SEARCH PEOPLE FROM SP
  // =========================================================
  const searchPeople = React.useCallback(
    async (value: string): Promise<PickerEntity[]> => {
      if (!value || value.length < 1) return [];

      const url = `${ctx.WebUrl}/_api/SP.UI.ApplicationPages.ClientPeoplePickerWebServiceInterface.ClientPeoplePickerSearchUser`;

      const body = {
        queryParams: {
          QueryString: value,
          MaximumEntitySuggestions: maxSuggestions,
          AllowEmailAddresses: true,
          AllowOnlyEmailAddresses: false,
          PrincipalType: principalType,
          PrincipalSource: 15,
        },
      };

      const response = await spHttpClient.post(url, spHttpClientConfig, {
        body: JSON.stringify(body),
      });
      const json = await response.json();

      if (!json || !json.d || !json.d.ClientPeoplePickerSearchUser) return [];

      const parsed = JSON.parse(json.d.ClientPeoplePickerSearchUser) as any[];

      const mapped: PickerEntity[] = parsed.map((p) => ({
        Key: String(p.EntityData?.SPUserID ?? p.Key),
        DisplayText: p.DisplayText,
        EntityType: "User",
        EntityData: {
          Email: p.EntityData?.Email,
          SPUserID: Number(p.EntityData?.SPUserID ?? p.Key),
        },
      }));

      return mapped;
    },
    [ctx.WebUrl, maxSuggestions, principalType, spHttpClient, spHttpClientConfig]
  );

  // =========================================================
  //             HYDRATE STARTER VALUE (Edit/View)
  // =========================================================
  React.useEffect(() => {
    if ((isEdit || isView) && starterValue) {
      const ids = Array.isArray(starterValue) ? starterValue : [starterValue];

      const loadValues = async () => {
        const allLoaded: PickerEntity[] = [];

        for (const id of ids) {
          const results = await searchPeople(String(id));
          const match = results.find((r) => Number(r.EntityData.SPUserID) === Number(id));
          if (match) allLoaded.push(match);
        }

        if (allLoaded.length > 0) {
          const newMap = new Map<string, PickerEntity>();
          const newKeys: string[] = [];

          for (const ent of allLoaded) {
            newMap.set(ent.Key, ent);
            newKeys.push(ent.Key);
          }

          setOptions(newMap);
          setSelectedOptions(newKeys);
          valueRef(allLoaded);
        }
      };

      loadValues();
    }
  }, [isEdit, isView, starterValue, searchPeople, valueRef]);

  // =========================================================
  //               RESOLVE SUGGESTIONS FOR PICKER
  // =========================================================
  const onResolveSuggestions = React.useCallback(
    async (filter: string, selected: ITag[]) => {
      if (!filter) return [];

      const results = await searchPeople(filter);

      const transformed: ITag[] = results.map((r) => {
        return {
          key: r.Key,
          name: r.DisplayText,
        };
      });

      return transformed.filter(
        (t) => !selected.some((s) => String(s.key) === String(t.key))
      );
    },
    [searchPeople]
  );

  // =========================================================
  //                  SINGLE + MULTI SELECT LOGIC
  // =========================================================
  const onOptionSelect = React.useCallback(
    (
      e: React.MouseEvent | React.KeyboardEvent,
      data: { selectedOptions: string[] }
    ) => {
      if (multiselect) {
        setSelectedOptions(data.selectedOptions);
      } else {
        const single = data.selectedOptions.length > 0 ? data.selectedOptions[0] : "";
        setSelectedOptions(single ? [single] : []);
      }

      setTouched(true);
    },
    [multiselect]
  );

  const handleInputChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const value = e.target.value;
    setQuery(value);
  };

  const handleBlur = (e: React.FocusEvent<HTMLInputElement>) => {
    setTouched(true);
    commitValue();
  };

  // =========================================================
  //                          RENDERING
  // =========================================================
  const selectedTags: ITag[] = selectedOptions
    .map((k) => options.get(k))
    .filter((p): p is PickerEntity => !!p)
    .map((p) => ({
      key: p.Key,
      name: p.DisplayText,
    }));

  return (
    <div className={styles.root}>
      <label>{displayName}</label>

      <TagPicker
        className={className}
        disabled={disabled || isView}
        selectedOptions={selectedOptions}
        onOptionSelect={onOptionSelect}
      >
        <TagPickerInput
          aria-label={displayName}
          value={query}
          placeholder={placeholder ?? "Search people..."}
          onChange={handleInputChange}
          onBlur={handleBlur}
        />
      </TagPicker>

      {description && <small>{description}</small>}
    </div>
  );
};

export default PeoplePicker;








