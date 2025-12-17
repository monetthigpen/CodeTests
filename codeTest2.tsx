const localStorageVar = `${conText.pageContext.web.title}.peoplePickerIDs`;
const storedRaw = localStorage.getItem(localStorageVar);

const checkSPUserIDStorage = storedRaw
  ? JSON.parse(storedRaw).find((x: any) => x?.Key === elm.Key)?.EntityData?.SPUserID
  : null;



