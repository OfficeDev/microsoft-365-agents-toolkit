// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { TurnContext, Request } from "@microsoft/agents-hosting";
import { CloudAdapter } from "@microsoft/agents-hosting";
import { CommandBot } from "./command.browser";
import { ConversationOptions } from "./interface";
import { NotificationBot } from "./notification.browser";
import { ErrorWithCode, ErrorCode, ErrorMessage } from "../core/errors";
import { formatString } from "../util/utils";

/**
 * Provide utilities for bot conversation, including:
 *   - handle command and response.
 *   - send notification to varies targets (e.g., member, group, channel).
 *
 * @remarks
 * Only work on server side.
 */
export class ConversationBot {
  /**
   * The bot adapter.
   *
   * @remarks
   * Only work on server side.
   */
  public readonly adapter: CloudAdapter;

  /**
   * The entrypoint of command and response.
   *
   * @remarks
   * Only work on server side.
   */
  public readonly command?: CommandBot;

  /**
   * The entrypoint of notification.
   *
   * @remarks
   * Only work on server side.
   */
  public readonly notification?: NotificationBot;

  /**
   * Create new instance of the `ConversationBot`.
   *
   * @param options - The initialize options.
   *
   * @remarks
   * Only work on server side.
   */
  public constructor(options: ConversationOptions) {
    throw new ErrorWithCode(
      formatString(ErrorMessage.BrowserRuntimeNotSupported, "ConversationBot"),
      ErrorCode.RuntimeNotSupported
    );
  }

  /**
   * The request handler to integrate with web request.
   *
   * @param req - An incoming HTTP [Request](xref:botbuilder.Request).
   * @param res - The corresponding HTTP [Response](xref:botbuilder.Response).
   * @param logic - The additional function to handle bot context.
   *
   * @remarks
   * Only work on server side.
   */
  public requestHandler(
    req: Request,
    res: any,
    logic?: (context: TurnContext) => Promise<any>
  ): Promise<void> {
    return Promise.reject(
      new ErrorWithCode(
        formatString(ErrorMessage.BrowserRuntimeNotSupported, "ConversationBot"),
        ErrorCode.RuntimeNotSupported
      )
    );
  }
}
