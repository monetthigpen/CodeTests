const handleChange = React.useCallback(
  (items?: ITag[]) => {
    const nextTags = items ?? [];

    // User has interacted – stop auto-hydration
    setHasUserEdited(true);

    // Update local tag state for the TagPicker itself
    setSelectedTags(nextTags);

    // If there’s no parent onChange, we still want to save via GlobalFormData
    const entities: PickerEntity[] = [];

    // Build quick lookup from last resolved entities
    const resolvedByKey = new Map<string, PickerEntity>();
    for (const e of lastResolved) {
      const rawKey =
        e.Key ??
        e.EntityData2?.AccountName ??
        e.EntityData2?.Email ??
        e.DisplayText ??
        "";
      resolvedByKey.set(String(rawKey).toLowerCase(), e);
    }

    // Map ITag[] back into PickerEntity[]
    for (const t of nextTags) {
      const lk = String(t.key).toLowerCase();
      const match = resolvedByKey.get(lk);

      if (match) {
        entities.push(match);
        continue;
      }

      // Optional: support free-text if you have allowFreeText
      if (allowFreeText) {
        entities.push({
          Key: String(t.key),
          DisplayText: t.name,
          IsResolved: false,
          EntityType: "User",
          EntityData2: /@/.test(String(t.key))
            ? { Email: String(t.key) }
            : undefined,
        });
      }
    }

    // ----- SAVE TO GLOBAL FORMDATA FOR SUBMIT / REQUIRED CHECKS -----
    const ids = getUserIds(entities);           // numeric SPUserID[]
    const targetId = `${id}Id`;                 // match what your other components use

    if (multiselect) {
      // multi-person: array of ids or []
      ctx.GlobalFormData(targetId, ids.length === 0 ? [] : ids);
    } else {
      // single-person: single id or null
      ctx.GlobalFormData(targetId, ids.length === 0 ? null : ids[0]);
    }

    // ----- GLOBAL REQUIRED ERROR HANDLING -----
    if (requiredEffective) {
      const errMsg =
        (multiselect ? ids.length === 0 : ids[0] == null)
          ? "This field is required."
          : "";
      ctx.GlobalErrorHandle(targetId, errMsg);
    }

    // Still notify parent DynamicFormKS if it passed an onChange
    if (onChange) {
      onChange(entities);
    }
  },
  [
    id,
    multiselect,
    allowFreeText,
    lastResolved,
    onChange,
    ctx,
    requiredEffective,
  ]
);






