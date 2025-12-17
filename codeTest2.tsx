const storedRaw = localStorage.getItem(localStorageVar) ?? "[]";
const storedArr = JSON.parse(storedRaw) as any[];

const checkSPUserIDStorage =
  (storedArr.find((x) => x?.Key === elm.Key)?.EntityData?.SPUserID as string) ?? "";

  if (!elm?.Key) continue;



