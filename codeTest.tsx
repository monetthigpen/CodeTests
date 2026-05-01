import * as React from "react";
import { FormDisplayMode } from "@microsoft/sp-core-library";
import { FormCustomizerContext } from "@microsoft/sp-listview-extensibility";
import { validateProcessMap } from "./flow/scaffold/validateprocessmap";
import { processMap } from "./flow/scaffold/processMap";
import DynamicForm from "@spfx-monorepo/shared-library/dist/cjs/components/DynamicForm";
import mainResource from "@spfx-monorepo/shared-library/dist/cjs/Utils/mainResources.json";
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

interface IUrlObj {
  url: string;
  method: "GET";
  headers: Record<string, string>;
}

interface IDataSource {
  listName: string;
  dataType: string;
  listUrlId?: string;
  graphAPI: boolean;
  exclude?: string;
}

const validate = validateProcessMap(processMap);
console.log(validate);

const AuthorizedRequestor: React.FC<IAuthorizedRequestorProps> = (props) => {
  const [initialMsg, setInitialMsg] = React.useState<string>("");
  const [errorItems, setErrorItems] = React.useState<string>("");
  const [createOpenDb, setCreateOpenDb] = React.useState<any>({});
  const [final, setFinal] = React.useState<boolean>(false);
  const [sourceType, setSourceType] = React.useState<string>("");

  const makeFetch = React.useCallback((fetchObj: IUrlObj): Promise<any> => {
    return new Promise((resolve) => {
      fetch(fetchObj.url, {
        method: fetchObj.method,
        headers: fetchObj.headers
      })
        .then((res) => resolve(res))
        .catch((error) => reject(error));
    });
  }, []);

  React.useEffect(() => {
    (async (): Promise<void> => {
      try {
        const resources = mainResource.items.filter(
          (resource: any) =>
            resource["site-id"] ===
            props.context.pageContext.site.id.toString()
        );

        if (!resources || resources.length === 0) {
          setErrorItems("Missing main resources in this site; cannot proceed");
          setInitialMsg("");
          return;
        }

        setInitialMsg("Getting Resources...");

        const siteJson: IUrlObj = {
          url: `${props.context.pageContext.site.serverRelativeUrl}/${resources[0].siteDetailsJSON}`,
          method: "GET",
          headers: {
            Accept: "application/json;odata=verbose"
          }
        };

        const formJSON: IUrlObj = {
          url: `${props.context.pageContext.site.serverRelativeUrl}/${resources[0].formDetailsJSON.replace(
            "{listName}",
            props.context.list.title
          )}`,
          method: "GET",
          headers: {
            Accept: "application/json;odata=verbose"
          }
        };

        const siteJSONCheck = await makeFetch(siteJson);
        const formJSONCheck = await makeFetch(formJSON);

        if (siteJSONCheck.status !== 200 || formJSONCheck.status !== 200) {
          const err =
            siteJSONCheck.status !== 200
              ? "Missing Resource - Site Details JSON; cannot proceed"
              : "Missing Resource - Form Details JSON; cannot proceed";

          setErrorItems(err);
          setInitialMsg("");
          return;
        }

        const listDb = await formJSONCheck.json();

        setSourceType(listDb.sourceType || "");

        const SPUserListName =
          "SP_User_" +
          props.context.pageContext.site.serverRelativeUrl.replace(
            "/sites/",
            ""
          );

        const dataSources: IDataSource[] = [
          {
            listName: props.context.list.title,
            dataType: "COL",
            graphAPI: true
          },
          {
            listName: "SP_UserProperties",
            dataType: "OTHER",
            listUrlId: "/_api/SP.UserProfiles.PeopleManager/GetMyProperties",
            graphAPI: false
          },
          {
            listName: SPUserListName,
            dataType: "OTHER",
            listUrlId:
              "/_api/web/currentuser?$select=Groups&$expand=Groups",
            graphAPI: false
          }
        ];

        if (
          props.displayMode !== FormDisplayMode.New &&
          listDb.sourceType === "LIB"
        ) {
          dataSources.push({
            listName: `${props.context.list.title}ITEM`,
            dataType: "OTHER",
            listUrlId: `/_api/web/lists/getbytitle('${props.context.list.title}')/items(${props.context.itemId})?$select=*,FileLeafRef`,
            graphAPI: false,
            exclude: "YES"
          });
        }

        setCreateOpenDb({
          siteGUID: props.context.pageContext.site.id,
          xmlurl: `${resources[0].xmlFilePath}`,
          siteUrl: props.context.pageContext.site.serverRelativeUrl,
          siteJson: `${resources[0].siteDetailsJSON}`,
          listDbInfo: dataSources
        });

        console.log("Data Sources:", dataSources);
        console.log("Resources:", resources[0]);

        setInitialMsg("");
        setErrorItems("");
        setFinal(true);
      } catch (error: any) {
        setErrorItems(error?.message || error?.toString() || "Unknown error");
        setInitialMsg("");
      }
    })();
  }, [props.context, props.displayMode, makeFetch]);

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
          <DynamicForm
            alldbInfo={createOpenDb}
            context={props.context}
            displayMode={props.displayMode}
            contentTypeId={props.context.contentTypeId}
            baseSourceType={sourceType}
            attachment={true}
            SaveButton={SaveComponent}
            formRules={formRules}
            Header={HeaderComponent}
            onSave={props.onSave}
            onClose={props.onClose}
          />
        </div>
      )}
    </>
  );
};

export default AuthorizedRequestor;