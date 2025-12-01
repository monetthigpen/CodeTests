const validate = React.useCallback(
  (items?: ITag[]): string => {
    // Use the items passed in if provided, otherwise fall back to current state
    const current = items ?? selectedTags;

    if (requiredEffective && (current ?? []).length === 0) {
      return REQUIRED_MSG;
    }
    return "";
  },
  [requiredEffective, selectedTags]
);







