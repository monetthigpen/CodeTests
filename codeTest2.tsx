const display = (elm.DisplayText ?? "").replace(/'/g, "''");

item.push({
  id: GrphIndex++,
  method: "GET",
  url: `${context.pageContext.web.absoluteUrl}/_api/web/siteusers?$select=Id,Title,Email&$filter=Title eq '${display}'`
});









