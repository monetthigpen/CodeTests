const name = (elm.DisplayText ?? "").replace(/'/g, "''");

item.push({
  id: `${GrphIndex++}`,
  method: "GET",
  url: `/users?$filter=displayName eq '${name}'&$select=id,displayName,mail,userPrincipalName`
});










