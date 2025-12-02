

// --- People-picker specific state ---
const [touched, setTouched] = React.useState<boolean>(false);
const [errorMsg, setErrorMsg] = React.useState<string | undefined>();



/**
 * ITag[] -> PickerEntity[]
 * This is the section that was giving you trouble.
 */
const handleChange = React.useCallback(
  (items?: ITag[]) => {
    const next = items ?? [];

    // See exactly what TagPicker is giving us
    console.log("PeoplePicker onChange ITag[]:", next);

    setSelectedTags(next);
    setTouched(true);

    const err = validate(next);
    reportError(err);

    if (!onChange2) {
      return;
    }

    const result: PickerEntity[] = [];

    // Build quick lookup from last resolved entities using SAME key logic as toTag
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

    if (resolvedByKey.size > 0) {
      // Normal path: map each tag back to a fully-hydrated entity
      for (const t of next) {
        const lk = String(t.key).toLowerCase();
        const hit = resolvedByKey.get(lk);

        if (hit) {
          result.push(hit);
        } else {
          // Fallback entity if for some reason it wasn't in lastResolved
          result.push({
            Key: String(t.key),
            DisplayText: t.name,
            IsResolved: false,
            EntityType: "User",
            EntityData2: {
              Email: /@/.test(String(t.key)) ? String(t.key) : undefined,
            },
          });
        }
      }
    } else {
      // Safety net: if lastResolved is empty, still return something based only on ITag
      for (const t of next) {
        result.push({
          Key: String(t.key),
          DisplayText: t.name,
          IsResolved: false,
          EntityType: "User",
          EntityData2: {
            Email: /@/.test(String(t.key)) ? String(t.key) : undefined,
          },
        });
      }
    }

    console.log("PeoplePicker onChange -> PickerEntity[]:", result);
    onChange2(result);
  },
  [lastResolved, onChange2, reportError, validate]
);






