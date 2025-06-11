import { useState } from "react";
import { Button, Spinner } from "@fluentui/react-components";
import config from "./lib/config";
import { app } from "@microsoft/teams-js";
import { createNestablePublicClientApplication } from "@azure/msal-browser";
import { useData } from "./lib/useData";

export function NestedAppAuth() {
  const [needConsent, setNeedConsent] = useState(false);
  const getActiveAccount = async (msalClient) => {
    let account = msalClient.getActiveAccount();
    if (!account) {
      try {
        const context = await app.getContext();
        const accountFilter = {
          tenantId: context.user?.tenant?.id,
          homeAccountId: context.user?.id,
          loginHint: context.user?.loginHint,
        };
        account = msalClient.getAccount(accountFilter);
        if (account) {
          msalClient.setActiveAccount(account);
          return account;
        } else {
          throw new Error(`Got no account`);
        }
      } catch (error) {
        throw new Error(`Failed to get active account: ${error}`);
      }
    }
    return account;
  };
  const { loading, data, error, reload } = useData(async () => {
    await app.initialize();
    const msalConfig = {
      auth: {
        clientId: config.clientId,
        authority: `https://login.microsoftonline.com/${config.tenantId}`,
        supportsNestedAppAuth: true,
      },
    };
    const msalClient = await createNestablePublicClientApplication(msalConfig);
    if (needConsent) {
      const result = await msalClient.loginPopup({
        scopes: ["User.Read"],
      });
      const account = result.account;
      msalClient.setActiveAccount(account);
      setNeedConsent(false);
    }
    try {
      const account = await getActiveAccount(msalClient);
      const result = await msalClient.acquireTokenSilent({
        scopes: ["User.Read"],
        account: account,
      });
      return result;
    } catch (error) {
      setNeedConsent(true);
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
      {!loading && !!error && <div className="error fixed">{error.toString()}</div>}
    </div>
  );
}
