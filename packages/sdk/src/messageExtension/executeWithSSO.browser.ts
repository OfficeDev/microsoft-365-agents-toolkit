// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import { TurnContext } from "@microsoft/agents-hosting";
import { OnBehalfOfCredentialAuthConfig } from "../models/configuration";
import { MessageExtensionTokenResponse } from "./teamsMsgExtTokenResponse";
import { ErrorWithCode, ErrorMessage, ErrorCode } from "../core/errors";
import { formatString } from "../util/utils";

/**
 * Users execute query with SSO or Access Token.
 *
 * @deprecated This package will be deprecated by 2026-07. Please use [Microsoft 365 Agents SDK](https://www.npmjs.com/package/@microsoft/agents-hosting) instead.
 *
 * @remarks
 * Only works in in server side.
 */
export function handleMessageExtensionQueryWithSSO(
  context: TurnContext,
  config: OnBehalfOfCredentialAuthConfig,
  initiateLoginEndpoint: string,
  scopes: string | string[],
  logic: (token: MessageExtensionTokenResponse) => Promise<any>
) {
  throw new ErrorWithCode(
    formatString(ErrorMessage.BrowserRuntimeNotSupported, "queryWithToken in message extension"),
    ErrorCode.RuntimeNotSupported
  );
}
