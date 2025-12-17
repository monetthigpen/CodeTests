import { SPHttpClient } from "@microsoft/sp-http"; // ONLY if you don’t already have it

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

  // SharePoint REST batch-ish: just parallel GETs
  const urls: string[] = batchFlag
    ? (requestsOrUrl?.requests ?? []).map((r: any) => r.url)
    : [requestsOrUrl as string];

  const results: Array<{ email: string; spUserId: number }> = [];

  await Promise.all(
    urls.map(async (url) => {
      const res = await context.spHttpClient.get(url, SPHttpClient.configurations.v1);
      const json = await res.json();

      // This endpoint returns { Id: number, Email: string, ... }
      const spUserId = Number(json?.Id);
      const email = String(json?.Email ?? "").toLowerCase();

      if (!Number.isNaN(spUserId) && spUserId > 0) {
        results.push({ email, spUserId });
      }
    })
  );

  console.log("normalized results:", results);

  // write back to keyValues (match by email)
  for (const kv of keyValues) {
    const kvEmail = (kv.Email ?? "").toLowerCase();
    const match = results.find((r) => r.email === kvEmail);
    if (match) {
      kv.EntityData = { ...(kv.EntityData ?? {}), SPUserID: String(match.spUserId) };
    }
  }

  console.log("keyValues updated:", keyValues);

  localStorage.setItem(localStorageVar, JSON.stringify(keyValues));
  console.log("saved to localStorage:", localStorageVar);
};

url: `${context.pageContext.web.absoluteUrl}/_api/web/siteusers/getByEmail('${encodeURIComponent(elm.EntityData?.Email ?? "")}')?$select=Id,Email`


