// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { Data, useData } from "./useData";
import {
  ErrorWithCode,
  TeamsUserCredential,
  TeamsUserCredentialAuthConfig,
} from "@microsoft/teamsfx";
import { Client, GraphError } from "@microsoft/microsoft-graph-client";
import { TokenCredentialAuthenticationProvider } from "@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials";
import { useState } from "react";

type GraphOptionWithCredential = {
  scope?: string[];
  credential?: TeamsUserCredential;
};

/**
 * Helper function to call Microsoft Graph API with authentication.
 *
 * @param fetchGraphDataAsync - async function of how to call Graph API and fetch data.
 * @param options - Authentication configuration and OAuth resource scope.
 * @returns data, loading status, error and reload function
 *
 * @deprecated This package will be deprecated by 2026-07.
 * @public
 */
export function useGraphWithCredential<T>(
  fetchGraphDataAsync: (
    graph: Client,
    credential: TeamsUserCredential,
    scope: string[],
  ) => Promise<T>,
  options?: GraphOptionWithCredential,
): Data<T> {
  let credential: TeamsUserCredential;
  if (!options?.credential) {
    const authConfig: TeamsUserCredentialAuthConfig = {
      clientId: process.env.REACT_APP_CLIENT_ID,
      initiateLoginEndpoint: process.env.REACT_APP_START_LOGIN_PAGE_URL,
    };
    credential = new TeamsUserCredential(authConfig);
  } else {
    credential = options?.credential;
  }

  let scope: string[];
  if (!options?.scope) {
    scope = ["User.Read"];
  } else {
    scope = options.scope;
  }

  const [needConsent, setNeedConsent] = useState(false);
  const { data, error, loading, reload } = useData(async () => {
    if (needConsent) {
      try {
        await credential.login(scope);
        setNeedConsent(false);
        // Important: tokens are stored in sessionStorage, read more here: https://aka.ms/teamsfx-session-storage-notice
      } catch (err: unknown) {
        if (err instanceof ErrorWithCode && err.message?.includes("CancelledByUser")) {
          const helpLink = "https://aka.ms/teamsfx-auth-code-flow";
          err.message +=
            '\nIf you see "AADSTS50011: The reply URL specified in the request does not match the reply URLs configured for the application" ' +
            "in the popup window, you may be using unmatched version for TeamsFx SDK (version >= 0.5.0) and Teams Toolkit (version < 3.3.0) or " +
            `cli (version < 0.11.0). Please refer to the help link for how to fix the issue: ${helpLink}`;
        }
        throw err;
      }
    }
    try {
      // Create an instance of the TokenCredentialAuthenticationProvider by passing the tokenCredential instance and options to the constructor
      const authProvider = new TokenCredentialAuthenticationProvider(credential, { scopes: scope });
      const graph = Client.initWithMiddleware({
        authProvider: authProvider,
      });
      const graphData = await fetchGraphDataAsync(graph, credential, scope);
      return graphData;
    } catch (err: unknown) {
      if (err instanceof GraphError && err.code?.includes("UiRequiredError")) {
        // Silently fail for user didn't consent error
        setNeedConsent(true);
      } else {
        throw err;
      }
    }
  });
  return { data, error, loading, reload };
}
