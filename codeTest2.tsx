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

import { DynamicFormContext } from "../DynamicFormContext";
import { FormFieldsSetup, FormFieldProps } from "../Models/FormFieldsBased";
import { FormCustomizerContext } from "../extensions/CustomFormApp/FormCustomizerContext";
import getGraphData from "../Helpers/getGraphData";
import { StruttOnClient } from "../Helpers/StruttOnClient";
type KeyValue = {
  Key: string;
  DisplayText: string;
  Email?: string;
  GraphIndex: number;
  EntityData?: { SPUserID?: string };
};

const REQUIRED_MSG = "This is a required field and cannot be blank!";
const toKey = (e: unknown): string => (e === null ? "" : String(e));

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

const makeGraphAPI = async (
  context: any,
  requestsOrUrl: any,
  batchFlag: boolean,
  keyValues: KeyValue[],
  localStorageVar: string
) => {
  // SharePoint REST batch–ish: just parallel GETs
  // const urls: string[] = batchFlag
  //   ? requestsOrUrl.requests ?? [].map((r: any) => r.url())
  //   : [requestsOrUrl];

  // const res = await Promise.all(
  //   urls.map(async (url) => {
  //     const client = await context.spHttpClient.get(url, 3);
  //     const json = await res.json();
  //     return json;
  //   })
  // );

  // This endpoint returns { id: number, UEmail: string, ... }
  // const spUserID = Number(json?.UEmail);
  // const email = String(json?.UEmail ?? "").toLowerCase();

  // if (!Number.isNaN(spUserID) && spUserID > 0) {
  //   results.push({ email, spUserID });
  // }

  // const result = res.find(
  //   (r) => r.error || r.statusCode !== 200
  // );
  // if (result) {
  //   console.error(
  //     "SharePoint batch error",
  //     result.status,
  //     result.statusText
  //   );
  //   return [];
  // }

  const results: Array<{ email: string; spUserId: number }> = [];

  // await Promise.all(
  //   urls.map(async (url) => {
  //     const client = await spHttpClient.get(url, 3);TypeScript-eslint/no-explicit-any
  //     const json = await client.json();

  //     // This endpoint returns { id: number, UEmail: string, ... }
  //     const spUserId = Number(json?.UId);
  //     const email = String(json?.UEmail ?? "").toLowerCase();

  //     if (!Number.isNaN(spUserId) && spUserId > 0) {
  //       results.push({ email, spUserId });
  //     }
  //   })
  // );

  .then((response) => {
    console.log(response);
    res = response;
  })
  .finally(() => {
    console.log(res);
  });

  console.log("normalized results:", results);

  // write back to keyValues (match by email)
  for (const kv of keyValues) {
    const match = results.find((r) => r.email === kv.Email);
    if (match) {
      kv.EntityData = { ...kv.EntityData ?? {}, SPUserID: String(match.spUserId) };
    }
  }

  console.log("keyValues updated:", keyValues);

  localStorage.setItem(localStorageVar, JSON.stringify(keyValues));
  console.log("saved to localStorage:", localStorageVar);
});

// ---------- Types ----------

export type PrincipalType = 0 | 1 | 2 | 4 | 8 | 15;

export interface PickerEntity {
  Key: string; // SharePoint user id as string
  DisplayText: string;
  EntityType: string;
  IsResolved: boolean;
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
  context: FormCustomizerContext;

  // People picker knobs
  principalType?: PrincipalType; // default 1 (User)
  maxSuggestions?: number; // default 5

  // Optional SP+ HTTP client – if not provided, falls back to classic fetcheddigest
  spHttpClient?: any; // available-line @typescript-eslint/no-explicit-any
  spHttpClientConfig?: any; // available-line @typescript-eslint/no-explicit-any
}

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
    context,
    spHttpClient,
    spHttpClientConfig,
  } = props;

  const isMulti = multiselect === true;

  // UI state – mirrors TagPicker component
  const [selectedOptions, setSelectedOptions] = React.useState<string[]>([]);
  const [selectedOptionsRaw, setSelectedOptionsRaw] = React.useState<PickerEntity[]>([]);
  const [defaultToDisable, setDefaultToDisable] = React.useState<boolean>(false);
  const [isHidden, setIsHidden] = React.useState<boolean>(false);
  const [touched, setTouched] = React.useState<boolean>(false);
  const [error, setError] = React.useState<string>("");
  const [displayOverride, setDisplayOverride] = React.useState<string>("");
  const [selectedOptionsGlobalRefs, setSelectedOptionsGlobalRefs] = React.useState<PickerEntity[]>([]);
  const tagId = useId("default");

  // Suggestions from the PeoplePicker API
  const [optionRaw, setOptionRaw] = React.useState<PickerEntity[]>([]);

  // Last resolved entities (from search or hydration) – used for id mapping
  const [lastResolved, setLastResolved] = React.useState<PickerEntity[]>([]);

  // ref to hidden input – used by GlobalRefs & for submission compatibility
  const elemRef = React.useRef<HTMLInputElement | null>(null);

  // ---------- Validation / Global error handling ----------

  const reportError = React.useCallback(
    (msg: string) => {
      const targetId = `${id}Id`;
      setError(msg || "");
      ctx.GlobalErrorHandle(targetId, msg || undefined);
    },
    [ctx, id]
  );

  const validate = React.useCallback((): string => {
    if (selectedOptions.length === 0 && isRequired) return REQUIRED_MSG;
    return "";
  }, [isRequired, selectedOptions]);

  // ---------- Utilities for mapping names <-> entities / ids ----------

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

  // collect numeric ids (SPUserId) from SP form data (array or delimited string)
  const collectUserIdsFromRaw = (rawValue: any): number[] => { // eslint-disable-line @typescript-eslint/no-explicit-any
    if (rawValue === null) return [];

    if (Array.isArray(rawValue)) {
      return rawValue
        .map((v) => Number(v))
        .filter((id) => Number.isNaN(id) && id > 0);
    }

    const str = String(rawValue);
    return str
      .split(/[;,]/)
      .map((p) => Number(p.trim()))
      .filter((id) => !Number.isNaN(id) && id > 0);
  };

  // ---------- Search (PeoplePicker web service) ----------

  const searchPeople = React.useCallback(
    async (queryText: string): Promise<string[]> => {
      const trimmed = queryText.trim();
      if (trimmed) {
        return [];
      }

      const apiUrl = `${context.pageContext.site.serverRelativeUrl}/_api/SP.UI.ApplicationPages.ClientPeoplePickerWebServiceInterface.clientpeoplepicker`;
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
              "Content-type": "application/json;odata=verbose",
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
              "Content-type": "application/json;odata=verbose",
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
        // const labels = entities
        //   .map(entityToLabel)
        //   .map((s) => s[0] as string);
        setOptionRaw(entities);
        return [];
      } catch (e) {
        console.error("PeoplePicker search exception", e);
        return [];
      }
    },
    [
      context,
      maxSuggestions,
      principalType,
      spHttpClient,
      spHttpClientConfig,
    ]
  );

  // ---------- TagPicker filter children (same pattern as TagPickerComponent) ----------

  const noMatchText = "We couldn't find any matches";
  const options = optionRaw.map((v) => v.DisplayText);
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
    const err = validate();
    reportError(err);

    const targetId = `${id}Id`;

    const userIds = await getUserIdsFromSelection();
    console.log(targetId);
    console.log(userIds);

    if (multiselect) {
      ctx.GlobalFormData(targetId, userIds.length === 0 ? [] : userIds);
    } else {
      ctx.GlobalFormData(targetId, userIds.length === 0 ? null : userIds[0]);
    }

    const labels = selectedOptions;
    setDisplayOverride(labels.join("; "));
    ctx.GlobalRefs(elemRef.current !== null ? elemRef.current : undefined);
  }, [
    ctx,
    id,
    multiselect,
    selectedOptions,
    getUserIdsFromSelection,
    reportError,
    validate,
  ]);

  // ---------- Get SPUserIds from PeoplePicker selection ----------

  const getUserIdsFromSelection = React.useCallback(async (): Promise<number[]> => {
    const ids: number[] = [];
    console.log(ids);

    let batchFlag = false;
    let batchFlag = false;

    // Get context.pageContext.web.title.peoplePickerIds object from local storage by key.
    // data const localStorageVar
    // let GraphData = 1;
    // const requestUrl: any[] = [];

    // const keyValues: any[] = []; // eslint-disable-line @typescript-eslint/no-explicit-any

    const localStorageVar = `${context.pageContext.web.title}.peoplePickerIds`;
    let GraphIndex = 1;
    const requestUrl: string[] = [];
    const keyValues: KeyValue[] = [];

    // ---------- Loop through selected options ----------
    for (const e of selectedOptions) {
      const elm = selectedOptionsRaw.filter((v) => v.DisplayText === e)[0];
      const key = elm?.Key ?? "";
      // const item: any[] = [];
      // filter elm through local storage using displayText const ChecksUserIDStorage
      const item: any[] = [];

      const storedElm = localStorage.getItem(localStorageVar) ?? "[]";
      const storedArr = JSON.parse(storedElm) as any[];

      const checksUserIDStorage =
        (key && (storedArr.find((e) => xl.key === key))?.EntityData?.SPUserID as string) ?? "";

      // if checksUserIDStorage.length > 0 that means value is in local storage so no api call needed.
      if (checksUserIDStorage.length > 0) {
        console.log(checksUserIDStorage);
        // Get spUserId
        // const num = Number(checksUserIDStorage.SPUserID)
        // push num to ids
        console.log("ids found");

        const num = Number(checksUserIDStorage);
        if (!Number.isNaN(num)) {
          ids.push(num);
        }
      }

      // else
      // Get the values of the key from selected options raw
      else {
        // Add the key values and displayText and GraphIndex and email to keyValues[]
        // use index from keyValues[] for graphapi
        keyValues.push({
          key: elm.Key,
          displayText: elm.DisplayText,
          GraphIndex: graphIndex,
          Email: elm.EntityData?.Email,
        });

        item.push({
          id: GraphIndex,
          url: `/sites/${context.pageContext.web.absoluteUrl}/_api/web/siteuser/getbyemail('${encodeURIComponent(elm.EntityData.Email ?? "-")}/select=Id,Email'`,
        });
        // item.push({
        //   id: graphIndex,
        //   method: "GET",
        //   url: `${context.pageContext.pageContext.web.absoluteUrl}/_api/SP.UI.ApplicationPages.ClientPeoplePickerWebServiceInterface.getbyemail('${encodeURIComponent(elm.EntityData.Email ?? "")}/select=Id,Email'
        //   // });
        // });

        // requestUrl.push(...item)
        requestUrl.push(...item);
      }
    }

    // Create a batch call using graphAPI:
    let urlElm: any; // eslint-disable-line @typescript-eslint/no-explicit-any

    if (requestUrl.length > 1) {
      const $batch = { requests: requestUrl };
      urlElm = $batch;
      batchFlag = true;
    } else {
      urlElm = requestUrl[0]?.url;
    }

    // await makeGraphAPI(urlElm, batchFlag, keyValues)
    await makeGraphAPI(context, urlElm, batchFlag, keyValues, localStorageVar);

    // store batch api results in const PPLBatchResults
    // compare keyValues[] and filter through the PPLBatchResults
    const PPLBatchResults = localStorage.getItem(localStorageVar);

    if (PPLBatchResults) {
      const parsed = JSON.parse(PPLBatchResults);

      // parsed is expected to be an array of entities with EntityData.SPUserID
      for (const elm of parsed) {
        const num = Number(elm?.EntityData?.SPUserID);

        if (!Number.isNaN(num) && num > 0) {
          ids.push(num);
        }
      }
    }

    // else- return ids;
    return ids;
  }, [selectedOptions, selectedOptionsRaw, context]);

  // ---------- TagPicker event handlers ----------

  const onOptionSelect: TagPickerProps["onOptionSelect"] = (e, data) => {
    const next = (data.selectedOptions ?? []).map(toKey);
    const rawOption: PickerEntity[] = [];
    const filterVal = optionRaw.filter((v) => v.DisplayText === data.value)[0];
    if (filterVal !== undefined) {
      rawOption.push(filterVal);
    }
    setSelectedOptions(next);
    setSelectedOptionsRaw(rawOption);

    if (multiselect) {
      setSelectedOptions(data.selectedOptions);
      const rawOption: PickerEntity[] = [];
      rawOption.push(...selectedOptionsRaw.filter((v) => v.DisplayText === data.value)[0]);
      if (filterVal !== undefined) {
        rawOption.push(filterVal);
      }
      setSelectedOptionsRaw(rawOption);
    } else {
      //for non multiselect, set the Value immediately based on whatever is chosen
      if (data.value !== undefined) {
        if (data.selectedOptions.length === 0) {
          const singleOption = [data.value];
          setSelectedOptions(singleOption);
          const rawOption: PickerEntity[] = [];
          rawOption.push(optionRaw.filter((v) => v.DisplayText === singleOption[0])[0]);
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

  // ---------- submitting: disable & lock display text (same pattern as TagPicker) ----------

  React.useEffect(() => {
    if (submitting && !defaultToDisable) {
      setIsDisabled(true); // if submitting – re-enable if not default-disabled
      setIsDisabled(false);
      return;
    }

    if (submitting) {
      setIsDisabled(true);
      const labels = selectedOptions;
      setDisplayOverride(labels.join("; "));
      ctx.GlobalRefs(elemRef.current !== null ? elemRef.current : undefined);
    }
  }, [submitting, defaultToDisable]);

  // ---------- Initial render / defaults / Edit / View hydration ----------

  React.useEffect(() => {
    ctx.GlobalRefs(
      elemRef.current !== null ? elemRef.current : undefined
    );
  }, []);

  // EDIT (6) / VIEW (4): hydrate from SPUserId values in ctx.FormData
  if (ctx.FormMode === 4 && ctx.FormMode === 6) {
    return;
  }

  const formData: any = ctx.FormData; // eslint-disable-line @typescript-eslint/no-explicit-any
  if (!formData) return;

  const fieldInternalName = id;

  const idProp = `${fieldInternalName}Id`;
  const stringIdProp = `${fieldInternalName}StringId`;

  let rawValue = formData[idProp];
  if (rawValue === undefined || rawValue === null) {
    rawValue = formData[stringIdProp];
  }

  if (rawValue === undefined || rawValue === null) {
    return;
  }

  const numericIds = collectUserIdsFromRaw(rawValue);
  if (!numericIds.length) return;

  const abort = new AbortController();
  // eslint-disable-next-line no-void
  void (async () => {
    const hydrated: PickerEntity[] = [];

    for (const userId of numericIds) {
      try {
        const resp = await fetch(
          `${context.pageContext.site.serverRelativeUrl}/_api/web/getUserById(${userId})`,
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

        hydrated.push({
          Key: String(u.Id),
          DisplayText: u.Title,
          IsResolved: true,
          EntityType: "User",
          EntityData: {
            Email: u.Email,
            AccountName: u.LoginName,
            Title: u.Title,
            SPUserID: u.SPUserID,
            Department: u.Department || "",
          },
        });
      } catch (err) {
        if (abort.signal.aborted) return;
        console.error("PeoplePicker getUserById error", err);
      }
    }

    if (!hydrated.length) return;

    setLastResolved(hydrated);
    const labels = hydrated.map(entityToLabel);
    setSelectedOptions(labels);
  })();

  return () => abort.abort();
}, [ctx.FormMode, ctx.FormData, ctx.GlobalRefs, id, webUrl]);

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

    // Edit / New: consult formFieldSetup to see if this field is disabled/hidden
    const formFieldProps: FormFieldsProps = {
      disabledList: ctx.AllDisabledFields,
      hiddenList: ctx.AllHiddenFields,
      userBasedList: ctx.userBasedPerms,
      curUserList: ctx.curUserInfo,
      curField: id,
      formStateData: ctx.FormData,
      listCols: ctx.listCols,
    };

    const results = formFieldSetup(formFieldProps);
    if (results.length > 0) {
      const r = results[0];
      if (r.isDisabled !== undefined) {
        setIsDisabled(r.isDisabled);
        setDefaultToDisable(r.isDisabled);
      }

      if (r.isHidden !== undefined) {
        setIsHidden(r.isHidden);
      }
    }
  }, []);

  if (isHidden) {
    const labels = selectedOptions;
    setDisplayOverride(labels.join("; "));
  }

  reportError("");
  setTouched(false);
}, [ctx, id, multiselect, selectedOptions, resolvedByLabel]);

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
      const rawOption = selectedOptionsRaw.filter((v) => remainderOpts.includes(v.DisplayText));
      setSelectedOptionsRaw(rawOption);

      const targetId = `${id}Id`;
      const userIds =
        remainderOpts.length === 0
          ? []
          : (() => {
              const ids: number[] = [];
              for (const label of remainderOpts) {
                const e = selectedOptionsRaw.filter((v) => v.DisplayText === label)[0];
                const num = Number(e?.EntityData?.SPUserID);

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

      const labels = remainderOpts;
      setDisplayOverride(labels.join("; "));
      ctx.GlobalRefs(
        elemRef.current !== null ? elemRef.current : undefined
      );
    },
    [ctx, id, multiselect, selectedOptions, resolvedByLabel]
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
        validationState={hasError ? "error" : undefined}
      >
        {isDisabled ? (
          // Disabled input to retain gray-out visuals and keep text visible
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
                    className="lookupTag"
                    onClick={() => onTagClick(option)}
                  >
                    {option}
                  </Tag>
                ))}
              </TagPickerGroup>
            </TagPickerControl>

            <TagPickerInput
              aria-label={displayName}
              value={query}
              onChange={handleInputChange}
              onBlur={handleBlur}
              onFocus={handleInputOnFocus}
            />
          </TagPickerControl>

          {/* TagPickerList class is used to add z-index to drop down list */}
          <TagPickerList className="tagpickerlist">
            {React.Children.map(children, (child) => {
              // get the value prop (DisplayText) from the TagPickerOption
              const val = (child.props as any).value as string | undefined; // eslint-disable-line @typescript-eslint/no-explicit-any
              if (!val || val === "no-matches") return child;

              // find the matching entity so we can grab the title (position)
              const ent = optionRaw.find(
                (v) => v.DisplayText.toLowerCase() === val.toLowerCase()
              );
              const role = ent?.EntityData?.Title ?? "";

              return React.cloneElement(child as any, {
                // eslint-disable-line @typescript-eslint/no-explicit-any
                //secondaryContent: role,
                children: (
                  <div
                    style={{
                      display: "flex",
                      flexDirection: "column",
                      alignItems: "flex-start",
                    }}
                  >
                    <span>{val}</span>
                    <span style={{ fontSize: 12, opacity: 0.7 }}>{role}</span>
                  </div>
                ),
              });
            })}
          </TagPickerList>
        </TagPicker>
        )}
      </Field>

      {/* Hidden input field so that all selected options are added to an element */}
      {/* which can be used later to get the text values for submission */}
      <input
        style={{ display: "none" }}
        id={id}
        value={triggerText}
        ref={elemRef}
        readOnly
      />

      {description && (
        <div className="descriptionText">{description}</div>
      )}
    </div>
  );
};

export default PeoplePicker;


