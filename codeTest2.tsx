url: `${context.pageContext.web.absoluteUrl}/_api/web/siteusers/getByEmail('${elm.EntityData?.Email ?? ""}')?$select=Id`

kv.EntityData = {
  ...(kv.EntityData ?? {}),
  SPUserID: String(user.Id)
};