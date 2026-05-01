import * as React from "react";
import { FormDisplayMode } from "@microsoft/sp-core-library";
import { FormCustomizerContext } from "@microsoft/sp-listview-extensibility";
import { validateProcessMap } from "../flowScaffold/validateprocessmap";
import { processMap } from "../flowScaffold/processMap";
import { allDBInfo, ListDBInfo } from "@spfx-monorepo/shared-library";
import DynamicFormKS from "@spfx-monorepo/shared-library/dist/cjs/components/DynamicFormKS";
import mainResource from "@spfx-monorepo/shared-library/dist/cjs/Utilities/mainResources.json";
import { Spinner } from "@fluentui/react-components";
import SaveComponent from "./SaveComponent";
import formRules from "../utils/formRules.json";
import HeaderComponent from "./HeaderComponent";

export interface IAuthorizedRequestorProps {
  context: FormCustomizerContext;
  displayMode: FormDisplayMode;
  onSave: () => void;
  onClose: () => void;
}

interface UrlObj {
  url: string;
  method: "GET";
  headers: {};
}

const validate = validateProcessMap(processMap);
console.log(validate);

const AuthorizedRequestor: React.FC<IAuthorizedRequestorProps> = (props) => {
  const [initialMsg, setInitialMsg] = React.useState<string>("");
  const [errorItems, setErrorItems] = React.useState<string>("");
  const [createOpenDB, setCreateOpenDB] = React.useState<allDBInfo[]>([]);
  const [final, setFinal] = React.useState<boolean>(false);
  const [sourceType, setSourceType] = React.useState<string>("");

  const makeFetch = (fetchObj: UrlObj): Promise<any> => {
    return new Promise((resolve, reject) => {
      let result: any;
      fetch(fetchObj.url, {
        method: fetchObj.method,
        headers: fetchObj.headers
      })
        .then((res) => {
          result = res;
        })
        .catch((error) => {
          reject(error);
        })
        .finally(() => {
          resolve(result);
        });
    });
  };

  React.useEffect(() => {
    const resources = mainResource.items.filter(
      (t) => t["site-id"] === props.context.pageContext.site.id.toString()
    );

    setInitialMsg("Getting Resources...");

    if (resources.length > 0) {
      const siteJSON: UrlObj = {
        url: `${props.context.pageContext.site.serverRelativeUrl}/${resources[0].siteDetailsJSON}`,
        method: "GET",
        headers: {
          Accept: "application/json;odata=verbose"
        }
      };

      try {
        (async () => {
          const siteJSONCheck = await makeFetch(siteJSON);

          if (siteJSONCheck.status !== 200) {
            const err = "Missing Resource - Site Details JSON; cannot proceed";
            setInitialMsg("");
            setErrorItems(err);
          } else {
            setSourceType(formRules.sourceType);

            const SPUsrListName = `SP.User-${props.context.pageContext.site.serverRelativeUrl.replace(
              "/sites/",
              ""
            )}`;

            const dataSources: ListDBInfo[] = [
              {
                listName: props.context.list.title,
                dataType: "COL",
                graphAPI: true
              },
              {
                listName: "SP.UserProperties",
                dataType: "OTHER",
                listUrlAdd:
                  "/_api/SP.UserProfiles.PeopleManager/GetMyProperties",
                validFor: 30,
                graphAPI: false
              },
              {
                listName: SPUsrListName,
                dataType: "OTHER",
                listUrlAdd:
                  "/_api/web/currentuser?$select=,Groups&$expand=Groups",
                validFor: 7,
                graphAPI: false
              }
            ];

            if (props.displayMode !== 8 && formRules.sourceType === "LIB") {
              dataSources.push({
                listName: `${props.context.list.title}ITEM`,
                dataType: "OTHER",
                listGUID: props.context.list.guid,
                listUrlAdd: `/_api/web/lists/GetByTitle('${props.context.list.title}')/getItemById('${props.context.item.ID}')?$select=*,FileLeafRef`,
                graphAPI: false,
                exclude: "YES"
              });
            }

            setCreateOpenDB([
              {
                siteGUID: props.context.pageContext.site.id,
                xmlurl: `${resources[0].XMLFilePath}`,
                siteUrl: props.context.pageContext.site.serverRelativeUrl,
                siteJSON: `${resources[0].siteDetailsJSON}`,
                listDbInfo: dataSources
              }
            ]);

            setInitialMsg("");
            setErrorItems("");
            setFinal(true);
          }
        })()
          .then(() => {})
          .catch(() => {})
          .finally(() => {});
      } catch (error: any) {
        setErrorItems(error);
        setInitialMsg("");
      }
    } else {
      setErrorItems("Missing main resources in this Site; cannot proceed");
      setInitialMsg("");
    }
  }, []);

  return (
    <>
      {initialMsg && (
        <div className="spinner-container">
          <Spinner labelPosition="after" label={initialMsg} />
        </div>
      )}

      {errorItems && (
        <div>
          <p className="error-msg">{errorItems}</p>
        </div>
      )}

      {final && (
        <div>
          <DynamicFormKS
            alldbsInfo={createOpenDB}
            context={props.context}
            displayMode={props.displayMode}
            contentTypeId={props.context.contentType.id}
            baseSourceType={sourceType}
            attachment={true}
            SaveButton={SaveComponent}
            formRules={formRules}
            Header={<HeaderComponent />}
            label="Authorized Requestor"
            onSave={props.onSave}
            onClose={props.onClose}
          />
        </div>
      )}
    </>
  );
};

export default AuthorizedRequestor;