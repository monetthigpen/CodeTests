item.push({
  id: GrphIndex++,
  method: "GET",
  url: `/users/${encodeURIComponent(elm.EntityData?.Email)}?$select=id,mail,userPrincipalName`
});










