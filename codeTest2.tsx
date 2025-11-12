<PeoplePicker
  id={listColumns[i].name}
  displayName={listColumns[i].displayName}
  starterValue={starterVal}
  isRequired={listColumns[i].required}
  submitting={isSubmitting}
  multiselect={
    // exactly how your TagPicker decides it:
    listColumns[i]?.allowMultipleValues ??
    listColumns[i]?.multi ??
    listColumns[i]?.lookup?.allowMultipleValues ??
    false
  }
  placeholder={listColumns[i].description}
  description={listColumns[i].description}
  className="elementsWidth"
  spHttpClient={props.context.spHttpClient}
  spHttpClientConfig={SPHttpClient.configurations.v1}
/>
