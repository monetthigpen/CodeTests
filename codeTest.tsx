console.log(
  "Cost Center column:",
  normalizedCreateOpenDB?.[0]?.ListDBInfo?.[0]?.ResultsData?.listData?.[0]?.columns?.find(
    (c: any) => c.displayName === "Cost Center"
  )
);

console.log(
  "Business Area column:",
  normalizedCreateOpenDB?.[0]?.ListDBInfo?.[0]?.ResultsData?.listData?.[0]?.columns?.find(
    (c: any) => c.displayName === "Business Area"
  )
);


