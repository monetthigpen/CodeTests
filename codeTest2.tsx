const raw = localStorage.getItem(localStorageVar);

if (raw) {
  const cached = JSON.parse(raw) as any[];
  const found = cached.find(kv => kv?.Key === elm?.Key);

  const spId = found?.EntityData?.SPUserID;
  if (spId) {
    ids.push(Number(spId));
    continue; // skip API call for this person
  }
}



