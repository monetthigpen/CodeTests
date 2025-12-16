const commitValue = React.useCallback(async () => {
  const err = validate();
  reportError(err);

  const targetId = `${id}Id`;

  const userIds = await getUserIdsFromSelection();
  console.log(userIds);

  if (multiselect) {
    ctx.GlobalFormData(targetId, userIds.length === 0 ? [] : userIds);
  } else {
    ctx.GlobalFormData(targetId, userIds.length === 0 ? null : userIds[0]);
  }

  const labels = selectedOptions;
  setDisplayOverride(labels.join("; "));
  ctx.GlobalRefs(elemRef.current !== null ? elemRef.current : undefined);
}, [
  ctx,
  id,
  multiselect,
  selectedOptions,
  getUserIdsFromSelection,
  reportError,
  validate,
]);









