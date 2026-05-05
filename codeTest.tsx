const normalizedCreateOpenDB = createOpenDB?.map((db: any) => ({
  ...db,
  listData: db.listData?.map((contentType: any) => ({
    ...contentType,
    columns: contentType.columns?.map((col: any) => {
      if (col.type === "user") {
        return {
          ...col,
          TypeAsString: "User",
          FieldTypeKind: 20,
          TypeDisplayName: "Person or Group"
        };
      }

      return col;
    })
  }))
}));



alldbsInfo={normalizedCreateOpenDB}


<DynamicFormKS
  alldbsInfo={normalizedCreateOpenDB}
  context={props.context}
  displayMode={props.displayMode}
  contentTypeId={props.context.contentType.id}
  baseSourceType={sourceType}
  attachment={true}
  saveButton={SaveComponent}
  formRules={formRules}
  Header={<HeaderComponent />}
  label="Authorized Requestors"
  onSave={props.onSave}
  onClose={props.onClose}
/>