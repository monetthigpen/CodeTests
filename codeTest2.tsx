// ------- EDIT / VIEW FORM: hydrate PeoplePicker from ctx.FormData (SPUserID) -------
React.useEffect(() => {
  // Only run for EditForm(6) or ViewForm(4)
  if (!(ctx.FormMode === 4 || ctx.FormMode === 6)) {
    // NewForm (8) etc → do nothing, picker starts empty
    return;
  }

  // Don't hydrate again if we already resolved something
  if (lastResolved.length > 0) {
    return;
  }

  const fieldInternalName = id;

  const formData = ctx.FormData as any | undefined;
  if (!formData) {
    return;
  }

  // 🔹 NEW: look at <InternalName>, then <InternalName>Id, then <InternalName>StringId
  let rawValue: any = formData[fieldInternalName];

  if (rawValue === undefined) {
    const idProp = `${fieldInternalName}Id`;
    const stringIdProp = `${fieldInternalName}StringId`;

    rawValue = formData[idProp] ?? formData[stringIdProp];
  }

  // Nothing saved for this field
  if (rawValue === null || rawValue === undefined) {
    return;
  }

  // ---- Helper: normalize whatever SharePoint stored into numeric SPUserID[] ----
  const collectIds = (value: any): number[] => {
    if (value === null || value === undefined) return [];

    // Already an array (multi-value people / lookup field)
    if (Array.isArray(value)) {
      const ids: number[] = [];
      for (const v of value) {
        if (v && typeof v === "object" && "Id" in v) {
          ids.push(Number((v as any).Id));
        } else {
          ids.push(Number(v));
        }
      }
      return ids.filter((id) => !Number.isNaN(id));
    }

    // String / number – could be "738;#729" or "738,729"
    const str = String(value);
    const parts = str.split(/[;,#]/);
    return parts
      .map((p) => Number(p.trim()))
      .filter((id) => !Number.isNaN(id));
  };

  const numericIds = collectIds(rawValue);

  if (!numericIds.length) {
    return;
  }

  const abort = new AbortController();

  (async () => {
    const hydrated: PickerEntity[] = [];

    for (const userId of numericIds) {
      try {
        const resp = await fetch(
          `${webUrl}/_api/web/getUserById(${userId})`,
          {
            method: "GET",
            headers: {
              Accept: "application/json;odata=verbose"
            },
            signal: abort.signal
          }
        );

        if (!resp.ok) {
          console.warn(
            "PeoplePicker getUserById failed",
            userId,
            resp.status,
            resp.statusText
          );
          continue;
        }

        const json: any = await resp.json();
        const u = json.d;

        hydrated.push({
          Key: String(u.Id),
          DisplayText: u.Title,
          IsResolved: true,
          EntityType: "User",
          EntityData2: {
            Email: u.Email,
            AccountName: u.LoginName,
            Title: u.Title,
            Department2: u.Department || ""
          }
        });
      } catch (err) {
        if (abort.signal.aborted) {
          return;
        }
        console.error("PeoplePicker getUserById error", err);
      }
    }

    if (!hydrated.length) {
      return;
    }

    // Store resolved entities & show them in the picker
    setLastResolved(hydrated);

    const tags = hydrated.map(toTag);
    setSelectedTags(tags);

    if (onChange) {
      onChange(hydrated);
    }
  })();

  return () => abort.abort();
}, [ctx.FormMode, ctx.FormData, id, lastResolved.length, onChange, webUrl]);



