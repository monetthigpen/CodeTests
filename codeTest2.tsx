const handleChange = React.useCallback(
  (items?: ITag[]) => {
    const next = items ?? [];

    // update what TagPicker shows
    setSelectedTags(next);

    // ----- validate + send to GlobalErrorHandle -----
    const err =
      requiredEffective && next.length === 0
        ? "This field is required."
        : "";
    const targetId = `${id}Id`;
    ctx.GlobalErrorHandle(targetId, err || undefined);

    // If no parent onChange was passed, we're done
    if (!onChange) {
      return;
    }

    // ----- map ITag[] -> PickerEntity[] -----
    const result: PickerEntity[] = [];

    // build lookup from lastResolved using the SAME key logic as toTag()
    const resolvedByKey = new Map<string, PickerEntity>();
    for (const e of lastResolved) {
      const rawKey =
        e.Key ??
        e.EntityData2?.AccountName ??
        e.EntityData2?.Email ??
        e.DisplayText ??
        "";

      if (!rawKey) continue;
      resolvedByKey.set(String(rawKey).toLowerCase(), e);
    }

    // now map each selected tag back to its entity
    for (const t of next) {
      const lk = String(t.key).toLowerCase();
      const ent = resolvedByKey.get(lk);

      if (ent) {
        result.push(ent);
      } else {
        // safety fallback: still return something built only from the tag
        result.push({
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

    // finally notify your original consumer
    onChange(result);
  },
  [ctx, id, requiredEffective, lastResolved, onChange]
);







