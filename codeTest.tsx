const makeFetch = React.useCallback((fetchObj: IUrlObj): Promise<Response> => {
  return fetch(fetchObj.url, {
    method: fetchObj.method,
    headers: fetchObj.headers
  });
}, []);