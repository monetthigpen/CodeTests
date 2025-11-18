/* -------- EDIT FORM: hydrate starterValue that contains SPUserID(s) -------- */
React.useEffect(() => {

  // Only run if we have starter values but haven't resolved anything yet
  if (!normalizedStarter.length || lastResolved.length > 0) {
      return;
  }

  // Try to interpret starter keys as numeric SPUserID values
  const numericIds = normalizedStarter
      .map(t => Number(t.key))
      .filter(id => !Number.isNaN(id));

  if (!numericIds.length) {
      return;
  }

  const abort = new AbortController();

  (async () => {
      const hydrated: PickerEntity[] = [];

      for (const id of numericIds) {
          try {
              const resp = await fetch(
                  `${webUrl}/_api/web/getUserById(${id})`,
                  {
                      method: "GET",
                      headers: {
                          Accept: "application/json;odata=verbose",
                      },
                      signal: abort.signal,
                  }
              );

              if (!resp.ok) {
                  console.warn("PeoplePicker getUserById failed", id, resp.status, resp.statusText);
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
                      Department2: u.Department || "",
                  }
              });

          } catch (err) {
              if (abort.signal.aborted) return;
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

}, [normalizedStarter, lastResolved.length, onChange, webUrl]);


