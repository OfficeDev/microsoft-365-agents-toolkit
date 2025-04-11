/* This code sample provides a starter kit to implement server side logic for your App in JavaScript,
 * refer to https://docs.microsoft.com/en-us/azure/azure-functions/functions-reference for complete Azure Functions
 * developer guide.
 */

// Import polyfills for fetch required by msgraph-sdk-javascript.
require("isomorphic-fetch");
const teamsfxSdk = require("@microsoft/teamsfx");
const { Client } = require("@microsoft/microsoft-graph-client");
const { app } = require("@azure/functions");
const {
  TokenCredentialAuthenticationProvider,
} = require("@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials");
const config = require("../config");

/**
 * This function handles requests from teamsapp client.
 * The HTTP request should contain an SSO token queried from Teams in the header.
 *
 * This function initializes the teamsapp SDK with the configuration and calls these APIs:
 * - new OnBehalfOfUserCredential(ssoToken, authConfig)  - Construct OnBehalfOfUserCredential instance with the received SSO token and initialized configuration.
 * - getUserInfo() - Get the user's information from the received SSO token.
 *
 * The response contains multiple message blocks constructed into a JSON object, including:
 * - An echo of the request body.
 * - The display name encoded in the SSO token.
 * - Current user's Microsoft 365 profile if the user has consented.
 *
 * @param {Context} context - The Azure Functions context object.
 * @param {HttpRequest} req - The HTTP request.
 */
async function getUserProfile(req, context) {
  context.log("HTTP trigger function processed a request.");

  // Initialize response.
  const res = {
    status: 200,
    body: {},
  };

  // Put an echo into response body.
  res.body.receivedHTTPRequestBody = (await req.text()) || "";

  // Prepare access token.
  const ssoToken = req.headers.get("Authorization")?.replace("Bearer ", "").trim();
  if (!ssoToken) {
    return {
      status: 400,
      body: JSON.stringify({
        error: "No access token was found in request header.",
      }),
    };
  }

  // Construct TeamsFx using user identity.
  let credential;
  try {
    const authConfig = {
      authorityHost: config.authorityHost,
      tenantId: config.tenantId,
      clientId: config.clientId,
      clientSecret: config.clientSecret,
    };
    credential = new teamsfxSdk.OnBehalfOfUserCredential(ssoToken, authConfig);
  } catch (e) {
    context.error(e);
    return {
      status: 500,
      body: JSON.stringify({
        error:
          "Failed to construct OnBehalfOfUserCredential using your ssoToken. " +
          "Ensure your function app is configured with the right Microsoft Entra App registration.",
      }),
    };
  }

  // Query user's information from the access token.
  try {
    const currentUser = await credential.getUserInfo();
    if (currentUser && currentUser.displayName) {
      res.body.userInfoMessage = `User display name is ${currentUser.displayName}.`;
    } else {
      res.body.userInfoMessage = "No user information was found in access token.";
    }
  } catch (e) {
    context.error(e);
    return {
      status: 400,
      body: JSON.stringify({
        error: "Access token is invalid.",
      }),
    };
  }

  // Create a graph client to access user's Microsoft 365 data after user has consented.
  try {
    // Create an instance of the TokenCredentialAuthenticationProvider by passing the tokenCredential instance and options to the constructor
    const authProvider = new TokenCredentialAuthenticationProvider(credential, {
      scopes: ["https://graph.microsoft.com/.default"],
    });

    // Initialize Graph client instance with authProvider
    const graphClient = Client.initWithMiddleware({
      authProvider: authProvider,
    });

    res.body.graphClientMessage = await graphClient.api("/me").get();
  } catch (e) {
    context.error(e);
    return {
      status: 500,
      body: JSON.stringify({
        error:
          "Failed to retrieve user profile from Microsoft Graph. The application may not be authorized.",
      }),
    };
  }

  res.body = JSON.stringify(res.body);

  return res;
}

app.http("getUserProfile", {
  methods: ["GET", "POST"],
  authLevel: "anonymous",
  handler: getUserProfile,
});
