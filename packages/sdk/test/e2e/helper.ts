// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import * as msal from "@azure/msal-node";
import { JwtPayload } from "jwt-decode";
import * as dotenv from "dotenv";
const urljoin = require("url-join");

export function extractIntegrationEnvVariables() {
  dotenv.config();
  if (!process.env.SDK_INTEGRATION_TEST_ACCOUNT) {
    throw new Error("Please set env SDK_INTEGRATION_TEST_ACCOUNT");
  }
  const accountData = process.env.SDK_INTEGRATION_TEST_ACCOUNT.split(";");
  if (accountData.length === 2) {
    process.env.SDK_INTEGRATION_TEST_ACCOUNT_NAME = accountData[0];
    process.env.SDK_INTEGRATION_TEST_ACCOUNT_PASSWORD = accountData[1];
  }
  if (!process.env.SDK_INTEGRATION_TEST_AAD) {
    throw new Error("Please set env SDK_INTEGRATION_TEST_AAD");
  }
  const aadData = process.env.SDK_INTEGRATION_TEST_AAD.split(";");
  if (aadData.length === 6) {
    process.env.SDK_INTEGRATION_TEST_AAD_AUTHORITY_HOST = aadData[0];
    process.env.SDK_INTEGRATION_TEST_AAD_TENANT_ID = aadData[1];
    process.env.SDK_INTEGRATION_TEST_USER_OBJECT_ID = aadData[2];
    process.env.SDK_INTEGRATION_TEST_M365_AAD_CLIENT_ID = aadData[3];
    process.env.SDK_INTEGRATION_TEST_M365_AAD_CLIENT_SECRET = aadData[4];
    process.env.SDK_INTEGRATION_TEST_M365_AAD_CERTIFICATE_CONTENT = aadData[5];
  }
  if (!process.env.SDK_INTEGRATION_TEST_API_CERTPROVIDER) {
    throw new Error("Please set env SDK_INTEGRATION_TEST_API_CERTPROVIDER");
  }
}

/**
 * Get Access Token from a specific AAD app client id.
 * @param clientId - remote or local AAD App id.
 * @param userName - Azure account login name.
 * @param password - Azure account login password.
 * @param tenantId - AAD App tenant id.
 * @param scope - AAD App custom scopes to restrict access to data and functionality protected by the API.
 */
export async function getAccessToken(
  clientId: string,
  userName: string,
  password: string,
  tenantId: string,
  scope?: string
): Promise<string> {
  const defaultAuthorityHost = process.env.SDK_INTEGRATION_TEST_AAD_AUTHORITY_HOST;
  const msalConfig = {
    auth: {
      clientId: clientId,
      authority: urljoin(defaultAuthorityHost!, tenantId!),
    },
  };
  let scopes: string[];
  // this scope is required.
  if (scope) {
    scopes = [scope];
  } else {
    const defaultScope = `api://localhost:53000/${process.env.SDK_INTEGRATION_TEST_M365_AAD_CLIENT_ID}/access_as_user`;
    scopes = [defaultScope!];
  }
  const pca = new msal.PublicClientApplication(msalConfig);
  const usernamePasswordRequest = {
    scopes: scopes,
    username: userName,
    password: password,
  };
  const response = await pca.acquireTokenByUsernamePassword(usernamePasswordRequest);
  return response!.accessToken;
}

 
/**
 * process.env.SDK_INTEGRATION_TEST_M365_AAD_CLIENT_ID is the Test AAD app mocking Teams first party app.
 * This function mocks the sso token get from Teams
 * @returns SSO token got from mocked Teams
 */
export async function getSsoTokenFromTeams(): Promise<string> {
  const missingConfigurations: string[] = [];
  if (!process.env.SDK_INTEGRATION_TEST_M365_AAD_CLIENT_ID) {
    missingConfigurations.push("SDK_INTEGRATION_TEST_M365_AAD_CLIENT_ID");
  }
  if (!process.env.SDK_INTEGRATION_TEST_ACCOUNT_NAME) {
    missingConfigurations.push("SDK_INTEGRATION_TEST_ACCOUNT_NAME");
  }
  if (!process.env.SDK_INTEGRATION_TEST_ACCOUNT_PASSWORD) {
    missingConfigurations.push("SDK_INTEGRATION_TEST_ACCOUNT_PASSWORD");
  }
  if (!process.env.SDK_INTEGRATION_TEST_AAD_TENANT_ID) {
    missingConfigurations.push("SDK_INTEGRATION_TEST_AAD_TENANT_ID");
  }

  if (missingConfigurations.length != 0) {
    throw new Error("Environment variables are missing: " + missingConfigurations.join(", "));
  }
  return await getAccessToken(
    process.env.SDK_INTEGRATION_TEST_M365_AAD_CLIENT_ID!,
    process.env.SDK_INTEGRATION_TEST_ACCOUNT_NAME!,
    process.env.SDK_INTEGRATION_TEST_ACCOUNT_PASSWORD!,
    process.env.SDK_INTEGRATION_TEST_AAD_TENANT_ID!,
    `api://localhost:53000/${process.env.SDK_INTEGRATION_TEST_M365_AAD_CLIENT_ID}/access_as_user`
  );
}

/**
 * Convert one-line certificate content to original format
 * @param content - certificate content of one-line format
 * @returns
 */
export function convertCertificateContent(content: string): string {
  return content.replace(/\\n/g, "\n");
}

export interface AADJwtPayLoad extends JwtPayload {
  appid?: string;
  idtyp?: string;
  scp?: string;
  upn?: string;
}
