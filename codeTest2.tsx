const localStorageVar = `${conText.pageContext.web.title}.peoplePickerIDs`;
const cachedRaw = localStorage.getItem(localStorageVar);
const cachedItems: any[] = cachedRaw ? JSON.parse(cachedRaw) : [];


const elm = selectedOptionsRaw.find((v) => v.DisplayText === e);
if (!elm?.Key) continue;

// find cached record by Key (best match)
const cached = cachedItems.find((x) => x?.Key === elm.Key);

const spIdStr = cached?.EntityData?.SPUserID;
const spIdNum = spIdStr ? Number(spIdStr) : NaN;

if (!Number.isNaN(spIdNum)) {
  ids.push(spIdNum);
  continue; // ✅ skip API call because cache hit
}




