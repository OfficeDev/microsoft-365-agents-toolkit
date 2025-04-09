// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import { CloudAdapter } from "@microsoft/agents-hosting";
import { CardActionMiddleware } from "../conversation/middlewares/cardActionMiddleware";
import { CardActionOptions, TeamsFxAdaptiveCardActionHandler } from "../conversation/interface";
import { ErrorWithCode, ErrorCode, ErrorMessage } from "../core/errors";
import { formatString } from "../util/utils";

/**
 * A card action bot to respond to adaptive card universal actions.
 *
 * @remarks
 * Only work on server side.
 */
export class CardActionBot {
  private readonly adapter: CloudAdapter;
  private middleware: CardActionMiddleware;

  /**
   * Create a new instance of the `CardActionBot`.
   *
   * @param adapter - The bound `CloudAdapter`.
   * @param options - The initialize options.
   */
  constructor(adapter: CloudAdapter, options?: CardActionOptions) {
    throw new ErrorWithCode(
      formatString(ErrorMessage.BrowserRuntimeNotSupported, "CardActionBot"),
      ErrorCode.RuntimeNotSupported
    );
  }

  /**
   * Register a card action handler to the bot.
   *
   * @param actionHandler - A card action handler to be registered.
   *
   * @remarks
   * Only work on server side.
   */
  registerHandler(actionHandler: TeamsFxAdaptiveCardActionHandler) {
    return Promise.reject(
      new ErrorWithCode(
        formatString(ErrorMessage.BrowserRuntimeNotSupported, "CardActionBot"),
        ErrorCode.RuntimeNotSupported
      )
    );
  }

  /**
   * Register card action handlers to the bot.
   *
   * @param actionHandlers - A set of card action handlers to be registered.
   *
   * @remarks
   * Only work on server side.
   */
  registerHandlers(actionHandlers: TeamsFxAdaptiveCardActionHandler[]) {
    return Promise.reject(
      new ErrorWithCode(
        formatString(ErrorMessage.BrowserRuntimeNotSupported, "CardActionBot"),
        ErrorCode.RuntimeNotSupported
      )
    );
  }
}
