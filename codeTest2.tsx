

const checkSPUserIDStorage =
  (key && (storedArr.find((x) => x?.Key === key)?.EntityData?.SPUserID as string)) ?? "";

if (checkSPUserIDStorage.length > 0) {
  console.log(checkSPUserIDStorage);
  console.log("ids found");

  const num = Number(checkSPUserIDStorage);
  if (!Number.isNaN(num)) ids.push(num);
} else {
  // ✅ now it will actually reach here for “not in local storage”
  // your API lookup code
}



