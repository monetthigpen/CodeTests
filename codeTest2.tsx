const getUserIdsFromSelection = React.useCallback(async (): Promise<number[]> => {
  const ids: number[] = [];

  // Track Graph requests
  let GrphIndex = 1;
  let batchFlag = false;
  const requestUri: any[] = [];
  const keyValues: any[] = [];

  // Loop through selected options
  for (const e of selectedOptions) {

    // Find the matching entity from selectedOptionsRaw
    const elm = selectedOptionsRaw.filter(
      (v) => v.DisplayText === e
    )[0];

    if (!elm) {
      continue;
    }

    // ---------------------------------------------
    // Filter elm through local storage using displayText / Key
    // ---------------------------------------------
    const checkSPUserIDStorage = localStorage.getItem(elm.Key);
    console.log(checkSPUserIDStorage);

    // If checkSPUserIDStorage length > 0
    // that means value is in local storage so no API call needed
    if (checkSPUserIDStorage != null) {
      console.log("ids found");

      const num = Number(checkSPUserIDStorage);
      if (!Number.isNaN(num)) {
        ids.push(num);
      }
    }

    // ---------------------------------------------
    // Else: value not found in local storage
    // Get values from selected options raw
    // ---------------------------------------------
    else {

      // Get the values of the Key from selected options raw
      const key = elm.Key;
      const displayText = elm.DisplayText;
      const email = elm.EntityData?.Email;

      // Add the key values and displayText and GraphIndex and email to keyValues[]
      keyValues.push({
        Key: key,
        DisplayText: displayText,
        Email: email,
        GraphIndex: GrphIndex
      });

      // Use index from keyValues[] for graph API
      const item = {
        id: GrphIndex++,
        method: "GET",
        url: `${context.pageContext.site.id}/_api/web/siteusers/getbyemail('${email}')`
      };

      // Push graph request item
      requestUri.push(item);
    }
  }

  // ---------------------------------------------
  // If requestUri.length > 0
  // Create a batch call using graphAPI
  // ---------------------------------------------
  if (requestUri.length > 0) {

    let urlElm: any;

    if (requestUri.length > 1) {
      const $batch = {
        requests: requestUri
      };

      urlElm = $batch;
      batchFlag = true;
    }

    // Single request (no batch)
    else {
      urlElm = requestUri[0].url;
      batchFlag = false;
    }

    // Execute Graph API request
    await makeGraphAPI(urlElm, batchFlag, keyValues);
  }

  // ---------------------------------------------
  // Store batch API results
  // Compare keyValues[] and filter through the PPBatchResults
  // ---------------------------------------------
  const PPBatchResults = keyValues;

  for (const elm of PPBatchResults) {
    const num = Number(elm.EntityData?.SPUserID);

    if (!Number.isNaN(num) && num > 0) {
      ids.push(num);
    }
  }

  // ---------------------------------------------
  // Return collected SharePoint User IDs
  // ---------------------------------------------
  return ids;

}, [selectedOptions, selectedOptionsRaw]);









