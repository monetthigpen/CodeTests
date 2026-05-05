const normalizedCreateOpenDB = createOpenDB?.map((db: any) => ({
  ...db,
  ListDBInfo: db.ListDBInfo?.map((listInfo: any) => ({
    ...listInfo,
    ResultsData: {
      ...listInfo.ResultsData,
      listData: listInfo.ResultsData?.listData?.map((item: any) => ({
        ...item,
        columns: item.columns?.map((col: any) => {
          if (
            col.name === "Authorized_x0020_Requestor" ||
            col.displayName === "Authorized Requestor"
          ) {
            return {
              ...col,
              type: "user",
              personOrGroup: {
                ...col.personOrGroup,
                allowMultipleSelection: false
              }
            };
          }

          return col;
        })
      }))
    }
  }))
}));


console.log("Normalized Authorized Requestor column:",
  normalizedCreateOpenDB?.[0]?.ListDBInfo?.[0]?.ResultsData?.listData?.[0]?.columns?.find(
    (c: any) => c.displayName === "Authorized Requestor"
  )
);



