const justSelectedRef = React.useRef<boolean>(false);

// Mark that an option was just selected (to skip blur validation)
justSelectedRef.current = true;
setTimeout(() => { justSelectedRef.current = false; }, 100);

// Mark that an option was just selected (to skip blur validation)
justSelectedRef.current = true;
setTimeout(() => { justSelectedRef.current = false; }, 100);
        
const handleBlur = (): void => {
  // Delay to let onOptionSelect run first
  setTimeout(() => {
    if (justSelectedRef.current) {
      return;
    }
    
    setTouched(true);
    commitValue();
  }, 150);
  
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