const normalizeUserFields = (value: any): any => {
  if (Array.isArray(value)) {
    return value.map(normalizeUserFields);
  }

  if (value && typeof value === "object") {
    const isAuthorizedRequestor =
      value.name === "Authorized_x0020_Requestor" ||
      value.displayName === "Authorized Requestor";

    const normalized: any = {};

    Object.keys(value).forEach((key) => {
      normalized[key] = normalizeUserFields(value[key]);
    });

    if (isAuthorizedRequestor) {
      return {
        ...normalized,
        type: "user",
        personOrGroup: {
          ...normalized.personOrGroup,
          allowMultipleSelection: false
        }
      };
    }

    return normalized;
  }

  return value;
};

const normalizedCreateOpenDB = normalizeUserFields(createOpenDB);


