// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { AccessToken } from "@azure/identity";
import { TurnContext } from "@microsoft/agents-hosting";
import { MessagingExtensionResponse } from "@microsoft/agents-hosting-teams";
import { ActivityTypes, Activity } from "@microsoft/agents-activity";
import { parseJwt, getScopesArray, formatString } from "../util/utils";
import { MessageExtensionTokenResponse } from "./teamsMsgExtTokenResponse";
import { ErrorWithCode, ErrorCode, ErrorMessage } from "../core/errors";
import { OnBehalfOfCredentialAuthConfig } from "../models/configuration";
import { internalLogger } from "../util/logger";
import { OnBehalfOfUserCredential } from "../credential/onBehalfOfUserCredential";
/**
 * Retrieve the OAuth Sign in Link to use in the MessagingExtensionResult Suggested Actions.
 * This method only work on MessageExtension with Query now.
 *
 * @param {OnBehalfOfCredentialAuthConfig} authConfig - User custom the message extension authentication configuration.
 * @param {initiateLoginEndpoint} initiateLoginEndpoint - Login page for Teams to redirect to.
 * @param {string | string[]} scopes - The list of scopes for which the token will have access.
 *
 * @returns SignIn link SilentAuth CardAction with 200 status code.
 */
function getSignInResponseForMessageExtensionWithSilentAuthConfig(
  authConfig: OnBehalfOfCredentialAuthConfig,
  initiateLoginEndpoint: string,
  scopes: string | string[]
): any {
  const scopesArray = getScopesArray(scopes);
  const signInLink = `${initiateLoginEndpoint}?scope=${encodeURI(scopesArray.join(" "))}&clientId=${
    authConfig.clientId
  }&tenantId=${authConfig.tenantId}`;
  return {
    composeExtension: {
      type: "silentAuth",
      suggestedActions: {
        actions: [
          {
            type: "openUrl",
            value: signInLink,
            title: "Message Extension OAuth",
          },
        ],
      },
    },
  };
}

/**
 *  Retrieve the OAuth Sign in Link to use in the MessagingExtensionResult Suggested Actions.
 * This method just a workaround for link unfurling now.
 *
 * @param {OnBehalfOfCredentialAuthConfig} authConfig - User custom the message extension authentication configuration.
 * @param {initiateLoginEndpoint} initiateLoginEndpoint - Login page for Teams to redirect to.
 * @param {string | string[]} scopes - The list of scopes for which the token will have access.
 *
 * @returns SignIn link Auth CardAction with 200 status code.
 */
function getSignInResponseForMessageExtensionWithAuthConfig(
  authConfig: OnBehalfOfCredentialAuthConfig,
  initiateLoginEndpoint: string,
  scopes: string | string[]
): any {
  const scopesArray = getScopesArray(scopes);
  const signInLink = `${initiateLoginEndpoint}?scope=${encodeURI(scopesArray.join(" "))}&clientId=${
    authConfig.clientId
  }&tenantId=${authConfig.tenantId}`;
  return {
    composeExtension: {
      type: "auth",
      suggestedActions: {
        actions: [
          {
            type: "openUrl",
            value: signInLink,
            title: "Message Extension OAuth",
          },
        ],
      },
    },
  };
}

/**
 * execution in message extension with SSO token.
 *
 * @param {TurnContext} context - The context object for the current turn.
 * @param {OnBehalfOfCredentialAuthConfig} authConfig - User custom the message extension authentication configuration.
 * @param {initiateLoginEndpoint} initiateLoginEndpoint - Login page for Teams to redirect to.
 * @param {string[]} scopes - The list of scopes for which the token will have access.
 * @param {function} logic - Business logic when executing the query in message extension with SSO or access token.
 *
 * @throws {@link ErrorCode|InternalError} when failed to get access token with unknown error.
 * @throws {@link ErrorCode|TokenExpiredError} when SSO token has already expired.
 * @throws {@link ErrorCode|ServiceError} when failed to get access token from simple auth server.
 * @throws {@link ErrorCode|InvalidParameter} when scopes is not a valid string or string array.
 * @throws {@link ErrorCode|RuntimeNotSupported} when runtime is nodeJS.
 *
 * @returns A MessageExtension Response for the activity. If the logic not return any, return void instead.
 */
export async function executionWithTokenAndConfig(
  context: TurnContext,
  authConfig: OnBehalfOfCredentialAuthConfig,
  initiateLoginEndpoint: string,
  scopes: string | string[],
  logic?: (token: MessageExtensionTokenResponse) => Promise<any>
): Promise<MessagingExtensionResponse | void> {
  const valueObj = context.activity.value as any;
  if (!valueObj.authentication || !valueObj.authentication.token) {
    internalLogger.verbose("No AccessToken in request, return silentAuth for AccessToken");
    return getSignInResponseForMessageExtensionWithSilentAuthConfig(
      authConfig,
      initiateLoginEndpoint,
      scopes
    );
  }
  try {
    const credential = new OnBehalfOfUserCredential(valueObj.authentication.token, authConfig);
    const token: AccessToken | null = await credential.getToken(scopes);
    const ssoTokenExpiration: number = parseJwt(valueObj.authentication.token).exp;
    const tokenRes: MessageExtensionTokenResponse = {
      ssoToken: valueObj.authentication.token,
      ssoTokenExpiration: new Date(ssoTokenExpiration * 1000).toISOString(),
      token: token!.token,
      expiration: token!.expiresOnTimestamp.toString(),
      connectionName: "",
    };
    if (logic) {
      return await logic(tokenRes);
    }
  } catch (err) {
    if (
      err instanceof ErrorWithCode &&
      err.code === ErrorCode.UiRequiredError &&
      context.activity.name === "composeExtension/query"
    ) {
      internalLogger.verbose("User not consent yet, return 412 to user consent first.");
      const response = { status: 412 };
      const activity = Activity.fromObject({
        value: response,
        type: ActivityTypes.InvokeResponse,
      });
      await context.sendActivity(activity);
    } else if (
      err instanceof ErrorWithCode &&
      err.code === ErrorCode.UiRequiredError &&
      context.activity.name === "composeExtension/queryLink"
    ) {
      internalLogger.verbose("User not consent yet, return auth card for user login");
      const response = getSignInResponseForMessageExtensionWithAuthConfig(
        authConfig,
        initiateLoginEndpoint,
        scopes
      );
      await context.sendActivity(
        Activity.fromObject({
          value: { status: 200, body: response },
          type: ActivityTypes.InvokeResponse,
        })
      );
      return;
    }
    throw err;
  }
}

/**
 * Users execute query in message extension with SSO or access token.
 *
 * @param {TurnContext} context - The context object for the current turn.
 * @param {OnBehalfOfCredentialAuthConfig} config - User custom the message extension authentication configuration.
 * @param {initiateLoginEndpoint} initiateLoginEndpoint - Login page for Teams to redirect to.
 * @param {string| string[]} scopes - The list of scopes for which the token will have access.
 * @param {function} logic - Business logic when executing the query in message extension with SSO or access token.
 *
 * @throws {@link ErrorCode|InternalError} when User invoke not response to message extension query.
 * @throws {@link ErrorCode|InternalError} when failed to get access token with unknown error.
 * @throws {@link ErrorCode|TokenExpiredError} when SSO token has already expired.
 * @throws {@link ErrorCode|ServiceError} when failed to get access token from simple auth server.
 * @throws {@link ErrorCode|InvalidParameter} when scopes is not a valid string or string array.
 * @throws {@link ErrorCode|RuntimeNotSupported} when runtime is nodeJS.
 *
 * @returns A MessageExtension Response for the activity. If the logic not return any, return void instead.
 */
export async function handleMessageExtensionQueryWithSSO(
  context: TurnContext,
  config: OnBehalfOfCredentialAuthConfig,
  initiateLoginEndpoint: string,
  scopes: string | string[],
  logic: (token: MessageExtensionTokenResponse) => Promise<any>
) {
  if (context.activity.name != "composeExtension/query") {
    internalLogger.error(ErrorMessage.OnlySupportInQueryActivity);
    throw new ErrorWithCode(
      formatString(ErrorMessage.OnlySupportInQueryActivity),
      ErrorCode.FailedOperation
    );
  }
  return await executionWithTokenAndConfig(
    context,
    config ?? {},
    initiateLoginEndpoint,
    scopes,
    logic
  );
}

/**
 * Users execute link query in message extension with SSO or access token.
 *
 * @param {TurnContext} context - The context object for the current turn.
 * @param {OnBehalfOfCredentialAuthConfig} config - User custom the message extension authentication configuration.
 * @param {initiateLoginEndpoint} initiateLoginEndpoint - Login page for Teams to redirect to.
 * @param {string| string[]} scopes - The list of scopes for which the token will have access.
 * @param {function} logic - Business logic when executing the link query in message extension with SSO or access token.
 *
 * @throws {@link ErrorCode|InternalError} when User invoke not response to message extension link query.
 * @throws {@link ErrorCode|InternalError} when failed to get access token with unknown error.
 * @throws {@link ErrorCode|TokenExpiredError} when SSO token has already expired.
 * @throws {@link ErrorCode|ServiceError} when failed to get access token from simple auth server.
 * @throws {@link ErrorCode|InvalidParameter} when scopes is not a valid string or string array.
 * @throws {@link ErrorCode|RuntimeNotSupported} when runtime is nodeJS.
 *
 * @returns A MessageExtension Response for the activity. If the logic not return any, return void instead.
 */
export async function handleMessageExtensionLinkQueryWithSSO(
  context: TurnContext,
  config: OnBehalfOfCredentialAuthConfig,
  initiateLoginEndpoint: string,
  scopes: string | string[],
  logic: (token: MessageExtensionTokenResponse) => Promise<any>
) {
  if (context.activity.name != "composeExtension/queryLink") {
    internalLogger.error(ErrorMessage.OnlySupportInLinkQueryActivity);
    throw new ErrorWithCode(
      formatString(ErrorMessage.OnlySupportInLinkQueryActivity),
      ErrorCode.FailedOperation
    );
  }
  return await executionWithTokenAndConfig(
    context,
    config ?? {},
    initiateLoginEndpoint,
    scopes,
    logic
  );
}
