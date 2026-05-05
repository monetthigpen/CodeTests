const normalizedCreateOpenDB = createOpenDB?.map((db: any) => ({
  ...db,
  listData: db.listData?.map((item: any) => ({
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
}));

