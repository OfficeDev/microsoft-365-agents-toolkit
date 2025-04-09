// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { CloudAdapter } from "@microsoft/agents-hosting";
import {
  BotSsoConfig,
  CommandOptions,
  CardActionOptions,
  ConversationReferenceStore,
} from "../conversation/interface";

/**
 * Options to initialize {@link NotificationBot}.
 */
export interface NotificationOptions {
  /**
   * An optional input of bot app Id.
   *
   * @remarks
   * If `botAppId` is not provided, `process.env.BOT_ID` will be used by default.
   */
  botAppId?: string;
  /**
   * An optional store to persist bot notification target references.
   *
   * @remarks
   * If `store` is not provided, a default conversation reference store will be used,
   * which stores notification target references into:
   *  - `.notification.localstore.json` if running locally
   *  - `${process.env.TEMP}/.notification.localstore.json` if `process.env.RUNNING_ON_AZURE` is set to "1"
   *
   * It's recommended to use your own store for production environment.
   */
  store?: ConversationReferenceStore;
}

/**
 * Options to initialize {@link ConversationBot}
 */
export interface ConversationOptions {
  /**
   * The bot adapter. If not provided, a default adapter will be created:
   * - with `adapterConfig` as constructor parameter.
   * - with a default error handler that logs error to console, sends trace activity, and sends error message to user.
   *
   * @remarks
   * If neither `adapter` nor `adapterConfig` is provided, will use BOT_ID and BOT_PASSWORD from environment variables.
   */
  adapter?: CloudAdapter;

  /**
   * If `adapter` is not provided, this `adapterConfig` will be passed to the new `CloudAdapter` when created internally.
   *
   * @remarks
   * If neither `adapter` nor `adapterConfig` is provided, will use BOT_ID and BOT_PASSWORD from environment variables.
   */
  adapterConfig?: { [key: string]: unknown };

  /**
   * Configurations for sso command bot
   */
  ssoConfig?: BotSsoConfig;

  /**
   * The command part.
   */
  command?: CommandOptions & {
    /**
     * Whether to enable command or not.
     */
    enabled?: boolean;
  };

  /**
   * The notification part.
   */
  notification?: NotificationOptions & {
    /**
     * Whether to enable notification or not.
     */
    enabled?: boolean;
  };

  /**
   * The adaptive card action handler part.
   */
  cardAction?: CardActionOptions & {
    /**
     * Whether to enable adaptive card actions or not.
     */
    enabled?: boolean;
  };
}
