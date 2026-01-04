const justSelectedRef = React.useRef<boolean>(false);

// Mark that an option was just selected (to skip blur validation)
justSelectedRef.current = true;
setTimeout(() => { justSelectedRef.current = false; }, 100);

// Mark that an option was just selected (to skip blur validation)
justSelectedRef.current = true;
setTimeout(() => { justSelectedRef.current = false; }, 100);
        
const handleBlur = (): void => {
  // Skip validation if an option was just selected
  if (justSelectedRef.current) {
    setQuery("");
    setOptionsRaw([]);
    return;
  }
  
  setTouched(true);
  commitValue();
  setQuery("");
  setOptionsRaw([]);
};

const isFirstRender = React.useRef(true);

React.useEffect(() => {
  if (isFirstRender.current) {
    isFirstRender.current = false;
    return;
  }
  
  if (!isMulti && selectedOptionsRaw.length > 0) {
    commitValue();
  }
}, [selectedOptionsRaw]); // eslint-disable-line react-hooks/exhaustive-deps