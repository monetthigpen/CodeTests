const handleBlur = (): void => {
  setQuery("");
  setOptionsRaw([]);
  
  // For single select, useEffect handles commit after selection
  if (!isMulti) {
    return;
  }
  
  setTimeout(() => {
    if (justSelectedRef.current) {
      return;
    }
    
    setTouched(true);
    commitValue();
  }, 150);
};

React.useEffect(() => {
  if (isFirstRender.current) {
    isFirstRender.current = false;
    return;
  }
  
  // For single select, commit directly without validation
  if (!isMulti && selectedOptionsRaw.length > 0) {
    reportError("");
    const targetId = `${id}Id`;
    getUserIdsFromSelection().then(userIds => {
      ctx.GlobalFormData(targetId, userIds.length > 0 ? userIds[0] : null);
      setDisplayOverride(selectedOptions.join("; "));
    });
  }
}, [selectedOptionsRaw]); // eslint-disable-line react-hooks/exhaustive-deps
