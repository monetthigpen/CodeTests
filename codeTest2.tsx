// PeoplePickerComponent.tsx
import * as React from "react";
import {
  TagPicker,
  TagPickerList,
  TagPickerInput,
  TagPickerControl,
  TagPickerGroup,
  TagPickerOption,
  TagPickerProps,
  useTagPickerFilter,
  Tag,
  Field,
  Textarea,
  useId,
} from "@fluentui/react-components";

import { DynamicFormContext } from "./DynamicFormContext";
import { formFieldsSetup, FormFieldsProps } from "../Utils/formFieldBased";
import { FormCustomizerContext } from "@microsoft/sp-listview-extensibility";
import getGraphData from "../Utils/getGraphApi";

type KeyValue = {
  Key: string;
  DisplayText: string;
  Email?: string;
  GraphIndex: number;
  EntityData?: { SPUserID?: string };
};

const makeGraphAPI = async (
  context: any,
  requestsDrUrl: any,
  batchFlag: boolean,
  keyValues: KeyValue[],
  localStorageVar: string
): Promise<void> => {
  console.log("makeGraphAPI called");
  console.log("batchFlag:", batchFlag);
  console.log("requestsDrUrl:", requestsDrUrl);

  let res: any;

  await getGraphData(context, requestsDrUrl, batchFlag)
    .then((response) => {
      console.log("Raw response:", response);
      res = response;
    })
    .catch((error) => {
      console.error("GraphAPI error:", error);
    })
    .finally(() => {
      console.log("GraphAPI response:", res);
    });

  // Process response and update keyValues
  if (!batchFlag && res?.value?.length > 0) {
    // Single request - list items response
    const item = res.value[0];
    console.log("Item from response:", item);
    
    // id is directly on the item, not in fields
    const spUserId = item?.id;
    console.log("Extracted SPUserID:", spUserId);
    
    if (spUserId && keyValues.length > 0) {
      keyValues[0].EntityData = { SPUserID: String(spUserId) };
    }
  } else if (batchFlag && res?.responses) {
    // Batch request - match by GraphIndex
    for (const resp of res.responses) {
      console.log("Batch response item:", resp);
      if (resp.status === 200 && resp.body?.value?.length > 0) {
        const item = resp.body.value[0];
        console.log("Batch item:", item);
        // id is directly on the item, not in fields
        const spUserId = item?.id;
        console.log("Batch item SPUserID:", spUserId);
        
        // Find matching keyValue by GraphIndex (resp.id)
        const kv = keyValues.find(k => k.GraphIndex === Number(resp.id));
        if (kv && spUserId) {
          kv.EntityData = { SPUserID: String(spUserId) };
        }
      }
    }
  }

  console.log("keyValues updated:", keyValues);
  localStorage.setItem(localStorageVar, JSON.stringify(keyValues));
  console.log("saved to localStorage:", localStorageVar);
};

// ---------- Types ----------

export type PrincipalType = 0 | 1 | 2 | 4 | 8 | 15;

export interface PickerEntity {
  Key: string; // SharePoint user Id as string
  DisplayText: string;
  EntityType?: string;
  IsResolved?: boolean;
  EntityData?: {
    Email?: string;
    AccountName?: string;
    Title?: string;
    SPUserID?: string;
    Department?: string;
  };
}

export interface PeoplePickerProps {
  id: string;

  displayName?: string;
  className?: string;
  description?: string;
  placeholder?: string;

  isRequired?: boolean;
  submitting?: boolean;

  multiselect?: boolean;
  disabled?: boolean;
  conText: FormCustomizerContext;

  // People picker knobs
  principalType?: PrincipalType; // default 1 (User)
  maxSuggestions?: number; // default 5

  // Optional SPFx HTTP client - if not provided, falls back to classic fetch+digest
  spHttpClient?: any; // eslint-disable-line @typescript-eslint/no-explicit-any
  spHttpClientConfig?: any; // eslint-disable-line @typescript-eslint/no-explicit-any
}

// ---------- Constants & helpers ----------

const REQUIRED_MSG = "This is a required field and cannot be blank!";
const toKey = (k: unknown): string => (k === null ? '' : String(k));

// convert a PickerEntity into a simple display label
const entityToLabel = (e: PickerEntity): string => {
  return (
    e.DisplayText ||
    e.EntityData?.Title ||
    e.EntityData?.Email ||
    e.EntityData?.SPUserID ||
    e.EntityData?.AccountName ||
    e.Key
  );
};

// collect numeric ids (SPUserId) from SP form data (array or delimited string)
const collectUserIdsFromRaw = (rawValue: any): number[] => { // eslint-disable-line @typescript-eslint/no-explicit-any
  if (rawValue === null) return [];

  if (Array.isArray(rawValue)) {
    return rawValue
      .map((v) => Number(v))
      .filter((id) => !Number.isNaN(id) && id > 0);
  }

  const str = String(rawValue);
  return str
    .split(/[;,]/)
    .map((p) => Number(p.trim()))
    .filter((id) => !Number.isNaN(id) && id > 0);
};

// ---------- Component ----------

const PeoplePicker: React.FC<PeoplePickerProps> = (props) => {
  const ctx = React.useContext(DynamicFormContext);

  const {
    id,
    displayName,
    className,
    description,
    placeholder,
    isRequired,

    submitting,
    multiselect,
    disabled,
    principalType = 1,
    maxSuggestions = 5,
    conText,
    spHttpClient,
    spHttpClientConfig,
  } = props;

  const isMulti = multiselect === true;

  // UI state - mirrors TagPickerComponent
  const [query, setQuery] = React.useState<string>("");
  const [selectedOptions, setSelectedOptions] = React.useState<string[]>([]);
  const [isDisabled, setIsDisabled] = React.useState<boolean>(!!disabled);
  const [defaultDisable, setDefaultDisable] = React.useState<boolean>(false);
  const [isHidden, setIsHidden] = React.useState<boolean>(false);
  const [touched, setTouched] = React.useState<boolean>(false);
  const [error, setError] = React.useState<string>('');
  const [displayOverride, setDisplayOverride] = React.useState<string>("");
  const [selectedOptionsRaw, setSelectedOptionsRaw] = React.useState<PickerEntity[]>([]);
  const tagId = useId("default");

  // Suggestions from the PeoplePicker API
  const [optionRaw, setOptionsRaw] = React.useState<PickerEntity[]>([]);

  // Last resolved entities (from search or hydration) - used for Id mapping
  const [lastResolved, setLastResolved] = React.useState<PickerEntity[]>([]);

  // ref to hidden input - used by GlobalRefs & for submission compatibility
  const elemRef = React.useRef<HTMLInputElement | null>(null);

  // ---------- Validation / Global error handling ----------

  const reportError = React.useCallback(
    (msg: string) => {
      const targetId = `${id}Id`;
      setError(msg || '');
      ctx.GlobalErrorHandle(targetId, msg || undefined);
    },
    [ctx.GlobalErrorHandle, id]
  );

  const validate = React.useCallback((): string => {
    if (selectedOptions.length === 0 && isRequired) return REQUIRED_MSG;
    return "";
  }, [isRequired, selectedOptions]);

  // ---------- Utilities for mapping names <-> entities / Ids ----------

  const resolvedByLabel = React.useMemo(() => {
    const map = new Map<string, PickerEntity>();
    for (const e of lastResolved) {
      const label = entityToLabel(e);
      if (label) {
        map.set(label.toLowerCase(), e);
      }
    }
    return map;
  }, [lastResolved]);

  // -------------------- Get SPUserIDs from PeoplePicker selection --------------------
  const getUserIdsFromSelection = React.useCallback(async (): Promise<number[]> => {
    const ids: number[] = [];
    console.log("getUserIdsFromSelection called");
    console.log("selectedOptions:", selectedOptions);
    console.log("selectedOptionsRaw:", selectedOptionsRaw);

    let batchFlag = false;

    const localStorageVar = `${conText.pageContext.web.title}.peoplePickerIDs`;
    let GrphIndex = 1;
    const requestUri: any[] = []; // eslint-disable-line @typescript-eslint/no-explicit-any
    const keyValues: KeyValue[] = [];

    // ------------------ Loop through selected options ------------------
    for (const e of selectedOptions) {
      const elm = selectedOptionsRaw.filter((v) => v.DisplayText === e)[0];
      const key = elm?.Key ?? "";
      const item: any[] = []; // eslint-disable-line @typescript-eslint/no-explicit-any

      // Check localStorage for cached SPUserID
      const storedRaw = localStorage.getItem(localStorageVar) ?? "[]";
      const storedArr = JSON.parse(storedRaw) as any[];

      const checkSPUserIDStorage =
        (key && (storedArr.find((x) => x?.Key === key))?.EntityData?.SPUserID as string) ?? "";

      // if checkSPUserIDStorage.length > 0 that means value is in local storage so no api call needed.
      if (checkSPUserIDStorage !== null && checkSPUserIDStorage.length > 0) {
        console.log(checkSPUserIDStorage);
        console.log("ids found");

        const num = Number(checkSPUserIDStorage);
        if (!Number.isNaN(num)) {
          ids.push(num);
        }
      } else {
        // Get the values of the Key from Selected options raw
        // Add the key values and displayText and GraphIndex and email to keyValues[]
        // use index from keyValues[] for graphapi

        keyValues.push({
          Key: elm.Key,
          DisplayText: elm.DisplayText,
          GraphIndex: GrphIndex,
          Email: elm.EntityData?.Email,
        });

        item.push({
          id: GrphIndex++,
          method: "GET",
          url: `/sites/${conText.pageContext.site.id}/lists/fe8fcb98-439f-4f47-af7c-ce27c61d945a/items?$expand=fields&$filter=fields/Title eq '${elm.DisplayText}'`
        });

        // item.push({
        //   id: GrphIndex++,
        //   method: "GET",
        //   url: `${conText.pageContext.web.absoluteUrl}/_api/web/siteusers/getByEmail('${encodeURIComponent(elm.EntityData?.Email ?? "")}')`
        // });

        // requestUri.push(...item);
        requestUri.push(...item);
      }
    }

    console.log("IDs found so far:", ids);
    console.log("API requests needed:", requestUri.length);

    // if requestUri.length > 0 - process batch or single request
    if (requestUri.length > 0) {
      let urlElm: any; // eslint-disable-line @typescript-eslint/no-explicit-any

      if (requestUri.length > 1) {
        const $batch = { requests: requestUri };
        urlElm = $batch;
        batchFlag = true;
      } else {
        urlElm = requestUri[0].url;
      }

      await makeGraphAPI(conText, urlElm, batchFlag, keyValues, localStorageVar);

      // Retrieve updated results from localStorage
      const PPLBatchResults = localStorage.getItem(localStorageVar);

      if (PPLBatchResults) {
        const parsed = JSON.parse(PPLBatchResults);
        console.log("Parsed localStorage after API call:", parsed);

        for (const item of parsed) {
          const num = Number(item?.EntityData?.SPUserID);

          if (!Number.isNaN(num) && num > 0 && !ids.includes(num)) {
            ids.push(num);
          }
        }
      }
    }

    console.log("Final IDs to return:", ids);
    return ids;
  }, [selectedOptions, selectedOptionsRaw, conText]);

  // ---------- Search (PeoplePicker Web Service) ----------

  const searchPeople = React.useCallback(
    async (queryText: string): Promise<string[]> => {
      const trimmed = queryText.trim();
      if (!trimmed) {
        return [];
      }

      const apiUrl = `${conText.pageContext.site.serverRelativeUrl}/_api/SP.UI.ApplicationPages.ClientPeoplePickerWebServiceInterface.ClientPeoplePickerSearchUser`;
      const body = JSON.stringify({
        queryParams: {
          AllowEmailAddresses: true,
          AllowMultipleEntities: false,
          AllUrlZones: false,
          MaximumEntitySuggestions: maxSuggestions,
          PrincipalSource: 15,
          PrincipalType: 1,
          QueryString: trimmed,
        },
      });

      try {
        let resp: Response;

        if (spHttpClient && spHttpClientConfig) {
          resp = await spHttpClient.post(apiUrl, spHttpClientConfig, {
            headers: {
              Accept: "application/json;odata=verbose",
              "Content-Type": "application/json;odata=verbose",
              "odata-version": "3.0",
            },
            body,
          });
        } else {
          const digest =
            (document.getElementById(
              "__REQUESTDIGEST"
            ) as HTMLInputElement | null)?.value || "";

          resp = await fetch(apiUrl, {
            method: "POST",
            headers: {
              Accept: "application/json;odata=verbose",
              "Content-Type": "application/json;odata=verbose",
              "X-RequestDigest": digest,
              "odata-version": "3.0",
            },
            body,
            credentials: "same-origin",
          });
        }

        if (!resp.ok) {
          const txt = await resp.text().catch(() => "");
          console.error(
            "PeoplePicker search error",
            resp.status,
            resp.statusText,
            txt
          );
          return [];
        }

        const json: any = await resp.json(); // eslint-disable-line @typescript-eslint/no-explicit-any
        const raw = json.d?.ClientPeoplePickerSearchUser ?? "[]";
        const entities: PickerEntity[] = JSON.parse(raw);
        setOptionsRaw(entities);
        return [];
      } catch (e) {
        console.error("PeoplePicker search exception", e);
        return [];
      }
    },
    [
      isMulti,
      maxSuggestions,
      principalType,
      spHttpClient,
      spHttpClientConfig,
    ]
  );

  // ---------- TagPicker filter children (same pattern as TagPickerComponent) ----------

  const noMatchText = "We couldn't find any matches";
  const options = optionRaw.map(v => v.DisplayText);
  const children = useTagPickerFilter({
    query,
    options,
    noOptionsElement: query.length >= 3 && options.length === 0 ? (
      <TagPickerOption value="no-matches">{noMatchText}</TagPickerOption>
    ) : (
      <></>
    ),
    filter: (option: string) =>
      !selectedOptions.includes(option) &&
      option.toLowerCase().includes(query.toLowerCase()),
  });

  // ---------- Commit value to GlobalFormData ----------

  const commitValue = React.useCallback(async () => {
    console.log("commitValue called");
    const err = validate();
    reportError(err);

    const targetId = `${id}Id`;
    console.log("targetId:", targetId);

    const userIds = await getUserIdsFromSelection();
    console.log("userIds from getUserIdsFromSelection:", userIds);

    if (multiselect) {
      console.log("Setting GlobalFormData (multiselect):", targetId, userIds);
      ctx.GlobalFormData(targetId, userIds.length === 0 ? [] : userIds);
    } else {
      console.log("Setting GlobalFormData (single):", targetId, userIds.length === 0 ? null : userIds[0]);
      ctx.GlobalFormData(targetId, userIds.length === 0 ? null : userIds[0]);
    }

    const labels = selectedOptions;
    setDisplayOverride(labels.join("; "));
    ctx.GlobalRefs(elemRef.current !== null ? elemRef.current : undefined);
  }, [
    id,
    selectedOptions,
    getUserIdsFromSelection,
    validate,
    reportError,
    multiselect,
    ctx,
  ]);

  // ---------- TagPicker event handlers ----------

  const onOptionSelect: TagPickerProps["onOptionSelect"] = (e, data) => {
    const next = (data.selectedOptions ?? []).map(toKey);

    if (touched) reportError(isRequired && next.length === 0 ? REQUIRED_MSG : '');
    if (data.value === "no-matches") {
      return;
    }
    if (multiselect) {
      setSelectedOptions(data.selectedOptions);
      const rawOption: PickerEntity[] = [];
      rawOption.push(...selectedOptionsRaw);
      const filterVal = optionRaw.filter(v => v.DisplayText === data.value)[0];
      if (filterVal !== undefined) {
        rawOption.push(filterVal);
      }
      setSelectedOptionsRaw(rawOption);
    } else {
      // For non multiselect, set the Value immediately based on whatever is chosen
      if (data.value !== undefined) {
        if (data.selectedOptions.length !== 0) {
          const singleOption = [data.value];
          setSelectedOptions(singleOption);
          const rawOption: PickerEntity[] = [];
          rawOption.push(optionRaw.filter(v => v.DisplayText === singleOption[0])[0]);
          setSelectedOptionsRaw(rawOption);
        }
      }
    }

    setQuery("");
  };

  const handleInputChange = React.useCallback(
    async (ev: React.ChangeEvent<HTMLInputElement>) => {
      const val = ev.target.value;
      setQuery(val);
      if (val.length >= 3) {
        await searchPeople(val);
      } else if (val.length === 0) {
        setQuery("");
        setOptionsRaw([]);
      }
    },
    [searchPeople]
  );

  const handleBlur = (): void => {
    setTouched(true);
    commitValue();
    setQuery("");
    setOptionsRaw([]);
  };

  // ---------- Submitting: disable & lock display text (same pattern as TagPicker) ----------

  React.useEffect(() => {
    if (!submitting && !defaultDisable) {
      // Form no longer submitting - re-enable if not default-disabled
      setIsDisabled(false);
      return;
    }

    if (submitting) {
      setIsDisabled(true);
      const labels = selectedOptions;
      setDisplayOverride(labels.join("; "));
      const err =
        isRequired && selectedOptions.length === 0 ? REQUIRED_MSG : "";
      reportError(err);
    }
  }, [submitting, defaultDisable]);

  // ---------- Initial render / defaults / Edit / View hydration ----------

  React.useEffect(() => {
    ctx.GlobalRefs(
      elemRef.current !== null ? elemRef.current : undefined
    );

    // EDIT (6) / VIEW (4): hydrate from SPUserId values in ctx.FormData
    if (ctx.FormMode !== 4 && ctx.FormMode !== 6) {
      return;
    }

    const formData: any = ctx.FormData; // eslint-disable-line @typescript-eslint/no-explicit-any
    if (!formData) return;

    const fieldInternalName = id;

    const idProp = `${fieldInternalName}Id`;
    const stringIdProp = `${fieldInternalName}IdStringId`;

    let rawValue = formData[idProp];
    if (rawValue === undefined || rawValue === null) {
      rawValue = formData[stringIdProp];
    }
    if (rawValue === undefined || rawValue === null) {
      return;
    }

    const numericIds = collectUserIdsFromRaw(rawValue);
    if (!numericIds.length) return;

    // *** IMMEDIATELY send starter values to GlobalFormData on initialization ***
    const targetId = `${fieldInternalName}Id`;
    if (multiselect) {
      ctx.GlobalFormData(targetId, numericIds);
    } else {
      ctx.GlobalFormData(targetId, numericIds[0]);
    }
    console.log("PeoplePicker initialization - sent to GlobalFormData:", targetId, numericIds);

    const abort = new AbortController();
    const localStorageVar = `${conText.pageContext.web.title}.peoplePickerIDs`;

    // eslint-disable-next-line no-void
    void (async () => {
      const hydrated: PickerEntity[] = [];
      const keyValuesToStore: KeyValue[] = [];
      const requestUri: any[] = [];
      let GrphIndex = 1;

      for (const userId of numericIds) {
        try {
          const resp = await fetch(
            `${conText.pageContext.site.serverRelativeUrl}/_api/web/getUserById(${userId})`,
            {
              method: "GET",
              headers: {
                Accept: "application/json;odata=verbose",
              },
              signal: abort.signal,
            }
          );

          if (!resp.ok) {
            console.warn(
              "PeoplePicker getUserById failed",
              userId,
              resp.status,
              resp.statusText
            );
            continue;
          }

          const json: any = await resp.json(); // eslint-disable-line @typescript-eslint/no-explicit-any
          const u = json.d;

          const entity: PickerEntity = {
            Key: String(u.Id),
            DisplayText: u.Title,
            IsResolved: true,
            EntityType: "User",
            EntityData: {
              Email: u.Email,
              AccountName: u.LoginName,
              Title: u.Title,
              SPUserID: String(u.Id),
              Department: u.Department || "",
            },
          };

          hydrated.push(entity);

          // Build keyValues for localStorage (matching line 251 pattern)
          keyValuesToStore.push({
            Key: entity.Key,
            DisplayText: entity.DisplayText,
            GraphIndex: GrphIndex++,
            Email: entity.EntityData?.Email,
            EntityData: { SPUserID: String(u.Id) },
          });

        } catch (err) {
          if (abort.signal.aborted) return;
          console.error("PeoplePicker getUserById error", err);
        }
      }

      if (!hydrated.length) return;

      // Store hydrated values to localStorage for future use
      localStorage.setItem(localStorageVar, JSON.stringify(keyValuesToStore));
      console.log("Hydration: saved to localStorage:", localStorageVar, keyValuesToStore);

      setLastResolved(hydrated);
      setSelectedOptionsRaw(hydrated);
      const labels = hydrated.map(entityToLabel);
      setSelectedOptions(labels);

      // *** UPDATED: Send starter value to global data on initialization ***
      const targetId = `${id}Id`;
      const userIds = hydrated
        .map(e => Number(e.EntityData?.SPUserID))
        .filter(num => !Number.isNaN(num) && num > 0);

      if (multiselect) {
        ctx.GlobalFormData(targetId, userIds.length === 0 ? [] : userIds);
      } else {
        ctx.GlobalFormData(targetId, userIds.length === 0 ? null : userIds[0]);
      }

      // Set display override
      setDisplayOverride(labels.join("; "));
    })();

    return () => abort.abort();
  }, []);

  // ---------- Disable / hidden logic (same as TagPicker) ----------

  React.useEffect(() => {
    // Display form: always disabled, just show text
    if (ctx.FormMode === 4) {
      setIsDisabled(true);
      const labels = selectedOptions;
      setDisplayOverride(labels.join("; "));
      reportError("");
      setTouched(false);
      return;
    }

    // Edit / New: consult formFieldsSetup to see if this field is disabled/hidden
    const formFieldProps: FormFieldsProps = {
      disabledList: ctx.AllDisableFields,
      hiddenList: ctx.AllHiddenFields,
      userBasedList: ctx.userBasedPerms,
      curUserList: ctx.curUserInfo,
      curField: id,
      formStateData: ctx.FormData,
      listColumns: ctx.listCols,
    };

    const results = formFieldsSetup(formFieldProps);
    if (results.length > 0) {
      const r = results[0];
      if (r.isDisabled !== undefined) {
        setIsDisabled(r.isDisabled);
        setDefaultDisable(r.isDisabled);
      }
      if (r.isHidden !== undefined) {
        setIsHidden(r.isHidden);
      }
    }

    if (isDisabled) {
      const labels = selectedOptions;
      setDisplayOverride(labels.join("; "));
    }

    reportError("");
    setTouched(false);
  }, []);

  const handleInputOnFocus = (): void => {
    setOptionsRaw([]);
  };

  // ---------- Derived view values ----------

  const selectedLabels = selectedOptions;
  const joinedText = selectedLabels.join("; ");
  const visibleText = displayOverride || joinedText;
  const triggerText = visibleText || "";
  const triggerPlaceholder = triggerText || (placeholder || "");

  const hasError = !!error;
  const disabledClass = isDisabled ? "is-disabled" : "";
  const rootClassName = [className, disabledClass].filter(Boolean).join(" ");

  // remove one selected option (used when clicking an existing tag)
  const onTagClick = React.useCallback(
    (option: string): void => {
      const remainderOpts = selectedOptions.filter((o) => o !== option);
      setSelectedOptions(remainderOpts);
      const rawOption = selectedOptionsRaw.filter(v => remainderOpts.includes(v.DisplayText));
      setSelectedOptionsRaw(rawOption);
      const targetId = `${id}Id`;
      const userIds =
        remainderOpts.length === 0
          ? []
          : (() => {
              const ids: number[] = [];
              for (const label of remainderOpts) {
                const e = selectedOptionsRaw.filter(v => v.DisplayText === label)[0];
                if (!e) continue;
                const num = Number(e.EntityData?.SPUserID);
                if (!Number.isNaN(num) && num > 0) ids.push(num);
              }
              return ids;
            })();

      if (multiselect) {
        ctx.GlobalFormData(targetId, userIds.length === 0 ? [] : userIds);
      } else {
        ctx.GlobalFormData(
          targetId,
          userIds.length === 0 ? null : userIds[0]
        );
      }

      // Update localStorage when tag is removed
      const localStorageVar = `${conText.pageContext.web.title}.peoplePickerIDs`;
      const updatedKeyValues = rawOption.map((entity, index) => ({
        Key: entity.Key,
        DisplayText: entity.DisplayText,
        GraphIndex: index + 1,
        Email: entity.EntityData?.Email,
        EntityData: { SPUserID: entity.EntityData?.SPUserID },
      }));
      localStorage.setItem(localStorageVar, JSON.stringify(updatedKeyValues));

      const labels = remainderOpts;
      setDisplayOverride(labels.join("; "));
      ctx.GlobalRefs(
        elemRef.current !== null ? elemRef.current : undefined
      );
    },
    [ctx, id, multiselect, selectedOptions, selectedOptionsRaw, resolvedByLabel, conText]
  );

  // ---------- Render ----------

  return (
    <div
      style={{ display: isHidden ? "none" : "block" }}
      className="fieldClass"
      aria-disabled={isDisabled ? "true" : undefined}
      data-disabled={isDisabled ? "true" : undefined}
    >
      <Field
        label={displayName}
        id={tagId}
        {...(isRequired && { required: true })}
        validationMessage={hasError ? error : undefined}
        validationState={hasError ? 'error' : undefined}
      >
        {isDisabled ? (
          // Disabled Input to retain gray-out visuals and keep text visible
          <Textarea
            id={id}
            disabled
            value={triggerText}
            placeholder={triggerPlaceholder}
            className={rootClassName}
            aria-disabled="true"
            data-disabled="true"
          />
        ) : (
          <TagPicker
            size="medium"
            onOptionSelect={onOptionSelect}
            selectedOptions={selectedOptions}
            inline={true}
            positioning="below-end"
          >
            <TagPickerControl aria-label={displayName}>
              <TagPickerGroup aria-label={displayName}>
                {selectedOptions.map((option) => (
                  <Tag
                    key={option}
                    shape="rounded"
                    value={option}
                    className="lookupTags"
                    onClick={() => onTagClick(option)}
                  >
                    {option}
                  </Tag>
                ))}
              </TagPickerGroup>
              <TagPickerInput
                aria-label={displayName}
                value={query}
                onChange={handleInputChange}
                onBlur={handleBlur}
                onFocus={handleInputOnFocus}
              />
            </TagPickerControl>

            {/* tagpickerList class is used to add z-index to drop down list */}
            <TagPickerList className="tagpickerList">
              {React.Children.map(children, (child) => {
                if (!React.isValidElement(child)) return child;

                // get the value prop (DisplayText) from the TagPickerOption
                const val = (child.props as any).value as string | undefined; // eslint-disable-line @typescript-eslint/no-explicit-any
                if (!val || val === "no-matches") return child;

                // find the matching entity so we can grab the Title (position)
                const ent = optionRaw.find(
                  (v) => v.DisplayText.toLowerCase() === val.toLowerCase()
                );
                const role = ent?.EntityData?.Title ?? "";

                return React.cloneElement(child as any, { // eslint-disable-line @typescript-eslint/no-explicit-any
                  children: (
                    <div
                      style={{
                        display: 'flex',
                        flexDirection: "column",
                        alignItems: "flex-start"
                      }}
                    >
                      <span>{val}</span>
                      <span style={{ fontSize: 12, opacity: 0.7 }}>{role}</span>
                    </div>
                  )
                });
              })}
            </TagPickerList>
          </TagPicker>
        )}

        {/* Hidden input field so that all selected options are added to an element
            which can be used later to get the text values for submission */}
        <input
          style={{ display: "none" }}
          id={id}
          value={triggerText}
          ref={elemRef}
          readOnly
        />
      </Field>

      {description && (
        <div className="descriptionText">{description}</div>
      )}
    </div>
  );
};

export default PeoplePicker;