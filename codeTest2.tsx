import getGraphData from "../Utils/getGraphApiIts"; // <-- update path to your file

type KeyValue = {
  Key: string;
  DisplayText: string;
  Email?: string;
  GraphIndex: number;
  EntityData?: { SPUserID?: string };
};

const makeGraphAPI = async (
  context: any,
  requestsOrUrl: any,
  batchFlag: boolean,
  keyValues: KeyValue[],
  localStorageVar: string
): Promise<void> => {
  console.log("makeGraphAPI called");
  console.log("batchFlag:", batchFlag);
  console.log("requestsOrUrl:", requestsOrUrl);

  // Call your existing Graph helper
  const graphResponse = await getGraphData(context, requestsOrUrl, batchFlag);
  console.log("graphResponse:", graphResponse);

  // ----------------------------
  // Normalize results into: [{ email, id }]
  // ----------------------------
  const results: Array<{ email: string; id: string }> = [];

  if (batchFlag) {
    // Graph $batch shape: { responses: [{ id, status, body: {...} }]}
    const responses = graphResponse?.responses ?? [];

    for (const r of responses) {
      const body = r?.body;

      // if the body is a collection (users?$filter...), Graph returns { value: [...] }
      const user =
        Array.isArray(body?.value) && body.value.length > 0 ? body.value[0] : body;

      const email = (user?.mail ?? user?.userPrincipalName ?? "").toLowerCase();
      const id = user?.id;

      if (email && id) {
        results.push({ email, id });
      }
    }
  } else {
    // Single call
    // Your helper returns response.value when it exists, else response
    const user =
      Array.isArray(graphResponse) && graphResponse.length > 0
        ? graphResponse[0]
        : graphResponse;

    const email = (user?.mail ?? user?.userPrincipalName ?? "").toLowerCase();
    const id = user?.id;

    if (email && id) {
      results.push({ email, id });
    }
  }

  console.log("normalized results:", results);

  // ----------------------------
  // Apply results back to keyValues[]
  // (SPUserID is just a string field — we’ll store the Graph user id)
  // ----------------------------
  for (const kv of keyValues) {
    const kvEmail = (kv.Email ?? "").toLowerCase();
    const match = results.find((r) => r.email === kvEmail);

    if (match) {
      kv.EntityData = { ...(kv.EntityData ?? {}), SPUserID: match.id };
    }
  }

  console.log("keyValues updated:", keyValues);

  // ----------------------------
  // Store in localStorage (since your other function reads from it)
  // ----------------------------
  localStorage.setItem(localStorageVar, JSON.stringify(keyValues));
  console.log("saved to localStorage:", localStorageVar);
};










