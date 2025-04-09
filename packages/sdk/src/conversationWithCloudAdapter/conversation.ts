// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import {
  CloudAdapter,
  TurnContext,
  Request,
  AuthConfiguration,
  loadAuthConfigFromEnv,
} from "@microsoft/agents-hosting";
import { CardActionBot } from "./cardAction";
import { CommandBot } from "./command";
import { BotSsoExecutionActivityHandler } from "../conversation/interface";
import { ConversationOptions } from "./interface";
import { NotificationBot } from "./notification";
import { DefaultBotSsoExecutionActivityHandler } from "../conversation/sso/defaultBotSsoExecutionActivityHandler";

/**
 * Provide utilities for bot conversation, including:
 *   - handle command and response.
 *   - send notification to varies targets (e.g., member, group, channel).
 *
 * @example
 * For command and response, you can register your commands through the constructor, or use the `registerCommand` and `registerCommands` API to add commands later.
 *
 * ```typescript
 * import { BotBuilderCloudAdapter } from "@microsoft/teamsfx";
 * import ConversationBot = BotBuilderCloudAdapter.ConversationBot;
 *
 * // register through constructor
 * const conversationBot = new ConversationBot({
 *   command: {
 *     enabled: true,
 *     commands: [ new HelloWorldCommandHandler() ],
 *   },
 * });
 *
 * // register through `register*` API
 * conversationBot.command.registerCommand(new HelpCommandHandler());
 * ```
 *
 * For notification, you can enable notification at initialization, then send notifications at any time.
 *
 * ```typescript
 * import { BotBuilderCloudAdapter } from "@microsoft/teamsfx";
 * import ConversationBot = BotBuilderCloudAdapter.ConversationBot;
 *
 * // enable through constructor
 * const conversationBot = new ConversationBot({
 *   notification: {
 *     enabled: true,
 *   },
 * });
 *
 * // get all bot installations and send message
 * for (const target of await conversationBot.notification.installations()) {
 *   await target.sendMessage("Hello Notification");
 * }
 *
 * // alternative - send message to all members
 * for (const target of await conversationBot.notification.installations()) {
 *   for (const member of await target.members()) {
 *     await member.sendMessage("Hello Notification");
 *   }
 * }
 * ```
 *
 * @remarks
 * Set `adapter` in {@link ConversationOptions} to use your own bot adapter.
 *
 * For command and response, ensure each command should ONLY be registered with the command once, otherwise it'll cause unexpected behavior if you register the same command more than once.
 *
 * For notification, set `notification.storage` in {@link ConversationOptions} to use your own storage implementation.
 */
export class ConversationBot {
  /**
   * The bot adapter.
   */
  public readonly adapter: CloudAdapter;

  /**
   * The entrypoint of command and response.
   */
  public readonly command?: CommandBot;

  /**
   * The entrypoint of notification.
   */
  public readonly notification?: NotificationBot;

  /**
   * The action handler used for adaptive card universal actions.
   */
  public readonly cardAction?: CardActionBot;

  /**
   * Create new instance of the `ConversationBot`.
   *
   * @remarks
   * It's recommended to create your own adapter and storage for production environment instead of the default one.
   *
   * @param options - The initialize options.
   */
  public constructor(options: ConversationOptions) {
    if (options.adapter) {
      this.adapter = options.adapter;
    } else {
      this.adapter = this.createDefaultAdapter(options.adapterConfig);
    }

    let ssoCommandActivityHandler: BotSsoExecutionActivityHandler | undefined;

    if (options?.ssoConfig) {
      if (options.ssoConfig.dialog?.CustomBotSsoExecutionActivityHandler) {
        ssoCommandActivityHandler =
          new options.ssoConfig.dialog.CustomBotSsoExecutionActivityHandler(options.ssoConfig);
      } else {
        ssoCommandActivityHandler = new DefaultBotSsoExecutionActivityHandler(options.ssoConfig);
      }
    }

    if (options.command?.enabled) {
      this.command = new CommandBot(
        this.adapter,
        options.command,
        ssoCommandActivityHandler,
        options.ssoConfig
      );
    }

    if (options.notification?.enabled) {
      this.notification = new NotificationBot(this.adapter, options.notification);
    }

    if (options.cardAction?.enabled) {
      this.cardAction = new CardActionBot(this.adapter, options.cardAction);
    }
  }

  private createDefaultAdapter(adapterConfig?: AuthConfiguration): CloudAdapter {
    const authConfig: AuthConfiguration = adapterConfig ?? loadAuthConfigFromEnv();
    const adapter = new CloudAdapter(authConfig);

    // the default error handler
    adapter.onTurnError = async (context, error) => {
      // This check writes out errors to console.
      console.error(`[onTurnError] unhandled error`, error);

      // Only send error message for user messages, not for other message types so the bot doesn't spam a channel or chat.
      if (context.activity.type === "message") {
        // Send a trace activity, which will be displayed in Bot Framework Emulator
        await context.sendTraceActivity(
          "OnTurnError Trace",
          error instanceof Error ? error.message : error,
          "https://www.botframework.com/schemas/error",
          "TurnError"
        );

        // Send a message to the user
        await context.sendActivity(`The bot encountered unhandled error: ${error.message}`);
        await context.sendActivity("To continue to run this bot, please fix the bot source code.");
      }
    };

    return adapter;
  }

  /**
   * The request handler to integrate with web request.
   *
   * @param req - An incoming HTTP [Request](xref:botbuilder.Request).
   * @param res - The corresponding HTTP [Response](xref:botbuilder.Response).
   * @param logic - The additional function to handle bot context.
   *
   * @example
   * For example, to use with Express:
   * ``` typescript
   * // The default/empty behavior
   * const expressApp = express();
   * expressApp.use(express.json());
   * expressApp.post("/api/notification", conversationBot.requestHandler);
   *
   * // Or, add your own logic
   * const expressApp = express();
   * expressApp.use(express.json());
   * expressApp.post("/api/notification", async (req, res) => {
   *   await conversationBot.requestHandler(req, res, async (context) => {
   *     // your-own-context-logic
   *   });
   * });
   * ```
   */
  public async requestHandler(
    req: Request,
    res: any,
    logic?: (context: TurnContext) => Promise<any>
  ): Promise<void> {
    if (logic === undefined) {
      // create empty logic
      logic = async () => {};
    }

    await this.adapter.process(req, res, logic);
  }
}
