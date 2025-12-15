// -------------------- Get SPUserIDs from PeoplePicker selection --------------------
const getUserIdsFromSelection = React.useCallback(async (): Promise<number[]> => {
  const ids: number[] = [];
  console.log(ids);

  // let batchFlag = false;
  let batchFlag = false;

  // Get context.pageContext.web.title.peoplePickerIDs object from local storage by key.
  // make const localStorageVar
  // let GrphIndex = 1;
  // const requestUri: any = [];
  const localStorageVar = `${context.pageContext.web.title}.peoplePickerIDs`;
  let GrphIndex = 1;
  const requestUri: any[] = []; // eslint-disable-line @typescript-eslint/no-explicit-any

  // -------------------- Loop through selected options --------------------
  for (const e of selectedOptions) {
    const elm = selectedOptionsRaw.filter((v) => v.DisplayText === e)[0];

    // const item = [];
    // filter elm through local storage using displayText const CheckSPUserIDStorage
    const item: any[] = []; // eslint-disable-line @typescript-eslint/no-explicit-any

    // filter elm through local storage using displayText const CheckSPUserIDStorage
    const checkSPUserIDStorage = localStorage.getItem(elm.Key);
    console.log(checkSPUserIDStorage);

    // If checkSPUserIDStorage.length > 0 that means value is in local storage so no api call needed.
    if (checkSPUserIDStorage != null && checkSPUserIDStorage.length > 0) {
      // Get SPUserId
      // const num = Number(checkSPUserIDStorage.SPUserID)
      // push num to ids
      console.log("ids found");

      const num = Number(checkSPUserIDStorage);
      if (!Number.isNaN(num)) {
        ids.push(num);
      }
    }

    // else
    // Get the values of the Key from Selected options raw
    else {
      // Add the key values and displayText and GraphIndex and email to keyValues[]
      // use index from keyValues[] for graphapi
      const keyValues: any[] = []; // eslint-disable-line @typescript-eslint/no-explicit-any

      keyValues.push({
        Key: elm.Key,
        DisplayText: elm.DisplayText,
        GraphIndex: GrphIndex,
        Email: elm.EntityData?.Email,
      });

      // item.push(
      //   {
      //     id: GrphIndex++,
      //     method: 'GET',
      //     url: `${context.pageContext.web.absoluteUrl}/_api/web/siteusers/getByEmail('${elm.EntityData?.Email}')`
      //   }
      // );
      item.push({
        id: GrphIndex++,
        method: "GET",
        url: `${context.pageContext.web.absoluteUrl}/_api/web/siteusers/getByEmail('${elm.EntityData?.Email}')`,
      });

      // requestUri.push(...item);
      requestUri.push(...item);
    }
  }

  // If requestUri.length > 0
  if (requestUri.length > 0) {
    // Create a batch call using graphAPI:
    // let urlElm;
    let urlElm: any; // eslint-disable-line @typescript-eslint/no-explicit-any

    // if (requestUri.length > 1) {
    if (requestUri.length > 1) {
      // const $batch = { requests: requestUri };
      // urlElm = $batch;
      // batchFlag = true;
      const $batch = { requests: requestUri };
      urlElm = $batch;
      batchFlag = true;
    }

    // else {
    else {
      // urlElm = requestUri[0].url;
      urlElm = requestUri[0].url;
    }

    // await makeGraphAPI(urlElm, batchFlag, keyValues)
    await makeGraphAPI(urlElm, batchFlag, localStorageVar); // NOTE: you still need to implement makeGraphAPI

    // store batch api results in const PPLBatchResults
    // compare keyValues[] and filter through the PPLBatchResults
    // const num = Number(elm.EntityData?.SPUserID)
    // if (!Number.isNaN(num) && num > 0) {
    //   ids.push(num);
    // }
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
  }

  // else- return ids;
  return ids;
}, [selectedOptions, selectedOptionsRaw, context]);









