// ----------------- Search (REST people API) -----------------
const searchPeople = React.useCallback(
  async (query: string): Promise<ITag[]> => {
    if (!query.trim()) {
      return [];
    }

    // NOTE: double underscore in __metadata is required
    const payload = {
      queryParams: {
        __metadata: {
          type: "SP.UI.ApplicationPages.ClientPeoplePickerQueryParameters",
        },
        AllowEmailAddresses: true,
        AllowMultipleEntities: isMulti,
        AllUrlZones: false,
        MaximumEntitySuggestions: maxSuggestions,
        PrincipalSource: 15,       // All sources
        PrincipalType: principalType, // 1 = Users
        QueryString: query,
      },
    };

    try {
      // Prefer SPHttpClient when available
      if (spHttpClient && spHttpClientConfig) {
        const resp = await spHttpClient.post(apiUrl, spHttpClientConfig, {
          headers: {
            Accept: "application/json;odata=verbose",
            "Content-Type": "application/json;odata=verbose",
          },
          body: JSON.stringify(payload),
        });

        if (!resp.ok) {
          const txt = await resp.text().catch(() => "");
          console.error(
            "PeoplePicker spHttpClient error",
            resp.status,
            resp.statusText,
            txt
          );
          return [];
        }

        const data: any = await resp.json();
        const raw = data.d?.ClientPeoplePickerSearchUser ?? "[]";
        const entities: PickerEntity[] = JSON.parse(raw);

        setLastResolved(entities);
        return entities.map(toTag);
      }

      // Fallback to classic fetch (same payload & headers)
      const resp = await fetch(apiUrl, {
        method: "POST",
        headers: {
          Accept: "application/json;odata=verbose",
          "Content-Type": "application/json;odata=verbose",
        },
        body: JSON.stringify(payload),
        credentials: "same-origin",
      });

      if (!resp.ok) {
        const txt = await resp.text().catch(() => "");
        console.error(
          "PeoplePicker fetch error",
          resp.status,
          resp.statusText,
          txt
        );
        return [];
      }

      const json: any = await resp.json();
      const raw = json.d?.ClientPeoplePickerSearchUser ?? "[]";
      const entities: PickerEntity[] = JSON.parse(raw);

      setLastResolved(entities);
      return entities.map(toTag);
    } catch (e) {
      console.error("PeoplePicker search exception", e);
      return [];
    }
  },
  [isMulti, maxSuggestions, principalType, spHttpClient, spHttpClientConfig, apiUrl]
);





