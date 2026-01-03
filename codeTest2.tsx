item.push({
          id: GrphIndex++,
          method: "GET",
          url: `${conText.pageContext.web.absoluteUrl}/_api/web/siteusers/getByEmail('${encodeURIComponent(elm.EntityData?.Email ?? "")}')`
        });



// Process response and update keyValues
  if (!batchFlag && res) {
    // Single request - siteusers response
    const item = res?.d || res;
    console.log("Item from response:", item);
    
    const spUserId = item?.Id;
    console.log("Extracted SPUserID:", spUserId);
    
    if (spUserId && keyValues.length > 0) {
      keyValues[0].EntityData = { SPUserID: String(spUserId) };
    }
  } else if (batchFlag && res?.responses) {
    // Batch request - match by GraphIndex
    for (const resp of res.responses) {
      console.log("Batch response item:", resp);
      if (resp.status === 200) {
        const item = resp.body?.d || resp.body;
        console.log("Batch item:", item);
        const spUserId = item?.Id;
        console.log("Batch item SPUserID:", spUserId);
        
        // Find matching keyValue by GraphIndex (resp.id)
        const kv = keyValues.find(k => k.GraphIndex === Number(resp.id));
        if (kv && spUserId) {
          kv.EntityData = { SPUserID: String(spUserId) };
        }
      }
    }
  }


