// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { CloudAdapter, ConversationReference, TurnContext } from "@microsoft/agents-hosting";
import { ChannelInfo, TeamDetails, TeamsChannelAccount } from "@microsoft/agents-hosting-teams";
import { ErrorWithCode, ErrorCode, ErrorMessage } from "../core/errors";
import { formatString } from "../util/utils";
import {
  MessageResponse,
  NotificationTarget,
  NotificationTargetType,
  ConversationReferenceStore,
} from "../conversation/interface";

/**
 * Send a plain text message to a notification target.
 *
 * @remarks
 * Only work on server side.
 *
 * @param target - The notification target.
 * @param text - The plain text message.
 * @param onError - An optional error handler that can catch exceptions during message sending.
 *
 * @returns A `Promise` representing the asynchronous operation.
 */
export function sendMessage(
  target: NotificationTarget,
  text: string,
  onError?: (context: TurnContext, error: Error) => Promise<void>
): Promise<void> {
  return Promise.reject(
    new ErrorWithCode(
      formatString(ErrorMessage.BrowserRuntimeNotSupported, "sendMessage"),
      ErrorCode.RuntimeNotSupported
    )
  );
}

/**
 * Send an adaptive card message to a notification target.
 *
 * @remarks
 * Only work on server side.
 *
 * @param target - The notification target.
 * @param card - The adaptive card raw JSON.
 * @param onError - An optional error handler that can catch exceptions during adaptive card sending.
 *
 * @returns A `Promise` representing the asynchronous operation.
 */
export function sendAdaptiveCard(
  target: NotificationTarget,
  card: unknown,
  onError?: (context: TurnContext, error: Error) => Promise<void>
): Promise<void> {
  return Promise.reject(
    new ErrorWithCode(
      formatString(ErrorMessage.BrowserRuntimeNotSupported, "sendAdaptiveCard"),
      ErrorCode.RuntimeNotSupported
    )
  );
}

/**
 * A {@link NotificationTarget} that represents a team channel.
 *
 * @remarks
 * Only work on server side.
 *
 * It's recommended to get channels from {@link TeamsBotInstallation.channels()}.
 */
export class Channel implements NotificationTarget {
  /**
   * The parent {@link TeamsBotInstallation} where this channel is created from.
   *
   * @remarks
   * Only work on server side.
   */
  public readonly parent: TeamsBotInstallation;

  /**
   * Detailed channel information.
   *
   * @remarks
   * Only work on server side.
   */
  public readonly info: ChannelInfo;

  /**
   * Notification target type. For channel it's always "Channel".
   *
   * @remarks
   * Only work on server side.
   */
  public readonly type: NotificationTargetType = NotificationTargetType.Channel;

  /**
   * Constructor.
   *
   * @remarks
   * Only work on server side.
   *
   * It's recommended to get channels from {@link TeamsBotInstallation.channels()}, instead of using this constructor.
   *
   * @param parent - The parent {@link TeamsBotInstallation} where this channel is created from.
   * @param info - Detailed channel information.
   */
  constructor(parent: TeamsBotInstallation, info: ChannelInfo) {
    throw new ErrorWithCode(
      formatString(ErrorMessage.BrowserRuntimeNotSupported, "Channel"),
      ErrorCode.RuntimeNotSupported
    );
  }

  /**
   * Send a plain text message.
   *
   * @remarks
   * Only work on server side.
   *
   * @param text - The plain text message.
   * @param onError - An optional error handler that can catch exceptions during message sending.
   *
   * @returns The response of sending message.
   */
  public sendMessage(
    text: string,
    onError?: (context: TurnContext, error: Error) => Promise<void>
  ): Promise<MessageResponse> {
    return Promise.reject(
      new ErrorWithCode(
        formatString(ErrorMessage.BrowserRuntimeNotSupported, "Channel"),
        ErrorCode.RuntimeNotSupported
      )
    );
  }

  /**
   * Send an adaptive card message.
   *
   * @remarks
   * Only work on server side.
   *
   * @param card - The adaptive card raw JSON.
   * @param onError - An optional error handler that can catch exceptions during adaptive card sending.
   *
   * @returns The response of sending adaptive card message.
   */
  public sendAdaptiveCard(
    card: unknown,
    onError?: (context: TurnContext, error: Error) => Promise<void>
  ): Promise<MessageResponse> {
    return Promise.reject(
      new ErrorWithCode(
        formatString(ErrorMessage.BrowserRuntimeNotSupported, "Channel"),
        ErrorCode.RuntimeNotSupported
      )
    );
  }
}

/**
 * A {@link NotificationTarget} that represents a team member.
 *
 * @remarks
 * Only work on server side.
 *
 * It's recommended to get members from {@link TeamsBotInstallation.members()}.
 */
export class Member implements NotificationTarget {
  /**
   * The parent {@link TeamsBotInstallation} where this member is created from.
   *
   * @remarks
   * Only work on server side.
   */
  public readonly parent: TeamsBotInstallation;

  /**
   * Detailed member account information.
   *
   * @remarks
   * Only work on server side.
   */
  public readonly account: TeamsChannelAccount;

  /**
   * Notification target type. For member it's always "Person".
   *
   * @remarks
   * Only work on server side.
   */
  public readonly type: NotificationTargetType = NotificationTargetType.Person;

  /**
   * Constructor.
   *
   * @remarks
   * Only work on server side.
   *
   * It's recommended to get members from {@link TeamsBotInstallation.members()}, instead of using this constructor.
   *
   * @param parent - The parent {@link TeamsBotInstallation} where this member is created from.
   * @param account - Detailed member account information.
   */
  constructor(parent: TeamsBotInstallation, account: TeamsChannelAccount) {
    throw new ErrorWithCode(
      formatString(ErrorMessage.BrowserRuntimeNotSupported, "Member"),
      ErrorCode.RuntimeNotSupported
    );
  }

  /**
   * Send a plain text message.
   *
   * @remarks
   * Only work on server side.
   *
   * @param text - The plain text message.
   * @param onError - An optional error handler that can catch exceptions during message sending.
   *
   * @returns The response of sending message.
   */
  public sendMessage(
    text: string,
    onError?: (context: TurnContext, error: Error) => Promise<void>
  ): Promise<MessageResponse> {
    return Promise.reject(
      new ErrorWithCode(
        formatString(ErrorMessage.BrowserRuntimeNotSupported, "Member"),
        ErrorCode.RuntimeNotSupported
      )
    );
  }

  /**
   * Send an adaptive card message.
   *
   * @remarks
   * Only work on server side.
   *
   * @param card - The adaptive card raw JSON.
   * @param onError - An optional error handler that can catch exceptions during adaptive card sending.
   *
   * @returns The response of sending adaptive card message.
   */
  public sendAdaptiveCard(
    card: unknown,
    onError?: (context: TurnContext, error: Error) => Promise<void>
  ): Promise<MessageResponse> {
    return Promise.reject(
      Promise.reject(
        new ErrorWithCode(
          formatString(ErrorMessage.BrowserRuntimeNotSupported, "Member"),
          ErrorCode.RuntimeNotSupported
        )
      )
    );
  }
}

/**
 * A {@link NotificationTarget} that represents a bot installation. Teams Bot could be installed into
 * - Personal chat
 * - Group chat
 * - Team (by default the `General` channel)
 *
 * @remarks
 * Only work on server side.
 *
 * It's recommended to get bot installations from {@link ConversationBot.installations()}.
 */
export class TeamsBotInstallation implements NotificationTarget {
  /**
   * The bound `CloudAdapter`.
   *
   * @remarks
   * Only work on server side.
   */
  public readonly adapter: CloudAdapter;

  /**
   * The bound `ConversationReference`.
   *
   * @remarks
   * Only work on server side.
   */
  public readonly conversationReference: Partial<ConversationReference>;

  /**
   * Notification target type.
   *
   * @remarks
   * Only work on server side.
   * - "Channel" means bot is installed into a team and notification will be sent to its "General" channel.
   * - "Group" means bot is installed into a group chat.
   * - "Person" means bot is installed into a personal scope and notification will be sent to personal chat.
   */
  public readonly type?: NotificationTargetType;

  /**
   * Constructor
   *
   * @remarks
   * Only work on server side.
   *
   * It's recommended to get bot installations from {@link ConversationBot.installations()}, instead of using this constructor.
   *
   * @param adapter - The bound `CloudAdapter`.
   * @param conversationReference - The bound `ConversationReference`.
   */
  constructor(adapter: CloudAdapter, conversationReference: Partial<ConversationReference>) {
    throw new ErrorWithCode(
      formatString(ErrorMessage.BrowserRuntimeNotSupported, "TeamsBotInstallation"),
      ErrorCode.RuntimeNotSupported
    );
  }

  /**
   * Send a plain text message.
   *
   * @remarks
   * Only work on server side.
   *
   * @param text - The plain text message.
   * @param onError - An optional error handler that can catch exceptions during message sending.
   *
   * @returns The response of sending message.
   */
  public sendMessage(
    text: string,
    onError?: (context: TurnContext, error: Error) => Promise<void>
  ): Promise<MessageResponse> {
    return Promise.reject(
      new ErrorWithCode(
        formatString(ErrorMessage.BrowserRuntimeNotSupported, "TeamsBotInstallation"),
        ErrorCode.RuntimeNotSupported
      )
    );
  }

  /**
   * Send an adaptive card message.
   *
   * @remarks
   * Only work on server side.
   *
   * @param card - The adaptive card raw JSON.
   * @param onError - An optional error handler that can catch exceptions during adaptive card sending.
   *
   * @returns The response of sending adaptive card message.
   */
  public sendAdaptiveCard(
    card: unknown,
    onError?: (context: TurnContext, error: Error) => Promise<void>
  ): Promise<MessageResponse> {
    return Promise.reject(
      new ErrorWithCode(
        formatString(ErrorMessage.BrowserRuntimeNotSupported, "TeamsBotInstallation"),
        ErrorCode.RuntimeNotSupported
      )
    );
  }

  /**
   * Get channels from this bot installation.
   *
   * @remarks
   * Only work on server side.
   *
   * @returns An array of channels if bot is installed into a team, otherwise returns an empty array.
   */
  public channels(): Promise<Channel[]> {
    return Promise.reject(
      new ErrorWithCode(
        formatString(ErrorMessage.BrowserRuntimeNotSupported, "TeamsBotInstallation"),
        ErrorCode.RuntimeNotSupported
      )
    );
  }

  /**
   * Get members from this bot installation.
   *
   * @remarks
   * Only work on server side.
   *
   * @returns An array of members from where the bot is installed.
   */
  public members(): Promise<Member[]> {
    return Promise.reject(
      new ErrorWithCode(
        formatString(ErrorMessage.BrowserRuntimeNotSupported, "TeamsBotInstallation"),
        ErrorCode.RuntimeNotSupported
      )
    );
  }

  /**
   * Get team details from this bot installation
   *
   * @returns The team details if bot is installed into a team, otherwise returns undefined.
   */
  public getTeamDetails(): Promise<TeamDetails | undefined> {
    return Promise.reject(
      new ErrorWithCode(
        formatString(ErrorMessage.BrowserRuntimeNotSupported, "TeamsBotInstallation"),
        ErrorCode.RuntimeNotSupported
      )
    );
  }
}

/**
 * Provide static utilities for bot notification.
 *
 * @remarks
 * Only work on server side.
 *
 * @example
 * Here's an example on how to send notification via Teams Bot.
 * ```typescript
 * // initialize (it's recommended to be called before handling any bot message)
 * const notificationBot = new NotificationBot(adapter);
 *
 * // get all bot installations and send message
 * for (const target of await notificationBot.installations()) {
 *   await target.sendMessage("Hello Notification");
 * }
 *
 * // alternative - send message to all members
 * for (const target of await notificationBot.installations()) {
 *   for (const member of await target.members()) {
 *     await member.sendMessage("Hello Notification");
 *   }
 * }
 * ```
 */
export class NotificationBot {
  private readonly conversationReferenceStore: ConversationReferenceStore;
  private readonly adapter: CloudAdapter;

  /**
   * Constructor of the notification bot.
   *
   * @remarks
   * Only work on server side.
   *
   * To ensure accuracy, it's recommended to initialize before handling any message.
   *
   * @param adapter - The bound `CloudAdapter`
   * @param options - The initialize options
   */
  public constructor(adapter: CloudAdapter, options?: NotificationOptions) {
    throw new ErrorWithCode(
      formatString(ErrorMessage.BrowserRuntimeNotSupported, "NotificationBot"),
      ErrorCode.RuntimeNotSupported
    );
  }

  /**
   * Get all targets where the bot is installed.
   *
   * @remarks
   * Only work on server side.
   *
   * The result is retrieving from the persisted storage.
   *
   * @returns An array of {@link TeamsBotInstallation}.
   */
  public static installations(): Promise<TeamsBotInstallation[]> {
    return Promise.reject(
      new ErrorWithCode(
        formatString(ErrorMessage.BrowserRuntimeNotSupported, "NotificationBot"),
        ErrorCode.RuntimeNotSupported
      )
    );
  }

  /**
   * Return the first {@link Member} where predicate is true, and undefined otherwise.
   *
   * @remarks
   * Only work on server side.
   *
   * @param predicate find calls predicate once for each member of the installation,
   * until it finds one where predicate returns true. If such a member is found,
   * find immediately returns that member. Otherwise, find returns undefined.
   * @param scope the scope to find members from the installations (personal chat, group chat, Teams channel).
   *
   * @returns The first {@link Member} where predicate is true, and undefined otherwise.
   */
  public findMember(
    predicate: (member: Member) => Promise<boolean>,
    scope?: SearchScope
  ): Promise<Member | undefined> {
    return Promise.reject(
      new ErrorWithCode(
        formatString(ErrorMessage.BrowserRuntimeNotSupported, "NotificationBot"),
        ErrorCode.RuntimeNotSupported
      )
    );
  }

  /**
   * Return the first {@link Channel} where predicate is true, and undefined otherwise.
   * (Ensure the bot app is installed into the `General` channel, otherwise undefined will be returned.)
   *
   * @remarks
   * Only work on server side.
   *
   * @param predicate - Find calls predicate once for each channel of the installation,
   * until it finds one where predicate returns true. If such a channel is found, find
   * immediately returns that channel. Otherwise, find returns undefined.
   *
   * @returns The first {@link Channel} where predicate is true, and `undefined` otherwise.
   */
  public findChannel(
    predicate: (channel: Channel, teamDetails: TeamDetails | undefined) => Promise<boolean>
  ): Promise<Channel | undefined> {
    return Promise.reject(
      new ErrorWithCode(
        formatString(ErrorMessage.BrowserRuntimeNotSupported, "NotificationBot"),
        ErrorCode.RuntimeNotSupported
      )
    );
  }

  /**
   * Return all {@link Member} where predicate is true, and empty array otherwise.
   *
   * @remarks
   * Only work on server side.
   *
   * @param predicate - Find calls predicate for each member of the installation.
   * @param scope - The scope to find members from the installations.
   * (personal chat, group chat, Teams channel).
   *
   * @returns An array of {@link Member} where predicate is true, and empty array otherwise.
   */
  public findAllMembers(
    predicate: (member: Member) => Promise<boolean>,
    scope?: SearchScope
  ): Promise<Member[]> {
    return Promise.reject(
      new ErrorWithCode(
        formatString(ErrorMessage.BrowserRuntimeNotSupported, "NotificationBot"),
        ErrorCode.RuntimeNotSupported
      )
    );
  }

  /**
   * Return all {@link Channel} where predicate is true, and empty array otherwise.
   * (Ensure the bot app is installed into the `General` channel, otherwise empty array will be returned.)
   *
   * @remarks
   * Only work on server side.
   *
   * @param predicate - Find calls predicate for each channel of the installation.
   *
   * @returns An array of {@link Channel} where predicate is true, and empty array otherwise.
   */
  public findAllChannels(
    predicate: (channel: Channel, teamDetails: TeamDetails | undefined) => Promise<boolean>
  ): Promise<Channel[]> {
    return Promise.reject(
      new ErrorWithCode(
        formatString(ErrorMessage.BrowserRuntimeNotSupported, "NotificationBot"),
        ErrorCode.RuntimeNotSupported
      )
    );
  }
}

/**
 * The search scope when calling {@link NotificationBot.findMember} and {@link NotificationBot.findAllMembers}.
 * The search scope is a flagged enum and it can be combined with `|`.
 * For example, to search from personal chat and group chat, use `SearchScope.Person | SearchScope.Group`.
 */
export enum SearchScope {
  /**
   * Search members from the installations in personal chat only.
   */
  Person = 1,

  /**
   * Search members from the installations in group chat only.
   */
  Group = 2,

  /**
   * Search members from the installations in Teams channel only.
   */
  Channel = 4,

  /**
   * Search members from all installations including personal chat, group chat and Teams channel.
   */
  All = Person | Group | Channel,
}
