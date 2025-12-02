React.useEffect(() => {

   // When submitting ends, re-enable if this field was not default-disabled
   if (!submitting && !defaultDisable) {
      setIsDisabled(false);
      return;
   }

   if (submitting) {
      // Disable while submitting
      setIsDisabled(true);

      // Build visible label text
      const labels = selectedOptions.map(o => o.DisplayText ?? o.text ?? '');
      setDisplayOverride(labels.join('; '));
   }

   // Validate after label update
   const next = selectedOptions ?? [];
   const isReq = props.isRequired ?? false;
   const msg = isReq && next.length === 0 ? REQUIRED_MSG : '';
   reportError(msg);

}, [submitting, defaultDisable]);  // ONLY these two deps








