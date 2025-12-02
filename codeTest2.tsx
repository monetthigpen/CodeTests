<TagPickerInput
  aria-label={displayName}
  value={query}
  placeholder={placeholder}
  onChange={(e: React.ChangeEvent<HTMLInputElement>) => {
    const value = e.target.value;
    setQuery(value);
    void searchPeople(value);
  }}
  onBlur={handleBlur}
/>








