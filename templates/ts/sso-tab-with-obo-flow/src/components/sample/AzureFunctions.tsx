import { useContext, useState } from "react";
import { Button, Spinner } from "@fluentui/react-components";
import { useData } from "@microsoft/teamsfx-react";
import * as axios from "axios";
import { BearerTokenAuthProvider, createApiClient, TeamsUserCredential } from "@microsoft/teamsfx";
import { TeamsFxContext } from "../Context";
import config from "./lib/config";

const functionName = config.apiName || "myFunc";

async function callFunction(teamsUserCredential: TeamsUserCredential) {
  try {
    const apiBaseUrl = config.apiEndpoint + "/api/";
    // createApiClient(...) creates an Axios instance which uses BearerTokenAuthProvider to inject token to request header
    const apiClient = createApiClient(
      apiBaseUrl,
      new BearerTokenAuthProvider(async () => (await teamsUserCredential.getToken(""))!.token)
    );
    const response = await apiClient.get(functionName);
    return response.data;
  } catch (err: unknown) {
    if (axios.default.isAxiosError(err)) {
      let funcErrorMsg = "";

      if (err?.response?.status === 404) {
        funcErrorMsg = `There may be a problem with the deployment of Azure Functions App, please deploy Azure Functions (Run command palette "Microsoft 365 Agents: Deploy") first before running this App`;
      } else if (err.message === "Network Error") {
        funcErrorMsg =
          "Cannot call Azure Functions due to network error, please check your network connection status and ";
        if (err.config?.url && err.config.url.indexOf("localhost") >= 0) {
          funcErrorMsg += `make sure to start Azure Functions locally (Run "npm run start" command inside api folder from terminal) first before running this App`;
        } else {
          funcErrorMsg += `make sure to provision and deploy Azure Functions (Run command palette "Microsoft 365 Agents: Provision" and "Microsoft 365 Agents: Deploy") first before running this App`;
        }
      } else {
        funcErrorMsg = err.message;
        if (err.response?.data?.error) {
          funcErrorMsg += ": " + err.response.data.error;
        }
      }

      throw new Error(funcErrorMsg);
    }
    throw err;
  }
}

export function AzureFunctions(props: { codePath?: string; docsUrl?: string }) {
  const [needConsent, setNeedConsent] = useState(false);
  const { codePath, docsUrl } = {
    codePath: `api/src/functions/${functionName}.ts`,
    docsUrl: "https://aka.ms/teamsfx-azure-functions",
    ...props,
  };
  const teamsUserCredential = useContext(TeamsFxContext).teamsUserCredential;
  const { loading, data, error, reload } = useData(async () => {
    if (!teamsUserCredential) {
      throw new Error("TeamsFx SDK is not initialized.");
    }
    if (needConsent) {
      await teamsUserCredential!.login(["User.Read"]);
      setNeedConsent(false);
    }
    try {
      const functionRes = await callFunction(teamsUserCredential);
      return functionRes;
    } catch (error: any) {
      if (error.message.includes("The application may not be authorized.")) {
        setNeedConsent(true);
      }
    }
  });
  return (
    <div>
      <h2>Call your Azure Functions</h2>
      <p>
        An Azure Functions app is running. Authorize this app and click below to call it for a
        response:
      </p>
      {!loading && (
        <Button appearance="primary" disabled={loading} onClick={reload}>
          Authorize and call Azure Functions
        </Button>
      )}
      {loading && (
        <pre className="fixed">
          <Spinner />
        </pre>
      )}
      {!loading && !!data && !error && <pre className="fixed">{JSON.stringify(data, null, 2)}</pre>}
      {!loading && !data && !error && <pre className="fixed"></pre>}
      {!loading && !!error && <div className="error fixed">{(error as any).toString()}</div>}
      <h4>How to edit the Azure Functions</h4>
      <p>
        See the code in <code>{codePath}</code> to add your business logic.
      </p>
      {!!docsUrl && (
        <p>
          For more information, see the{" "}
          <a href={docsUrl} target="_blank" rel="noreferrer">
            docs
          </a>
          .
        </p>
      )}
    </div>
  );
}
