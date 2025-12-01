// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

const { CardFactory } = require("@microsoft/agents-hosting");
const { TeamsInfo } = require("@microsoft/agents-hosting-teams");
const { ActivityTypes } = require("@microsoft/agents-activity");
const { NotificationMiddleware, getKey } = require("./middleware");

function getTeamsBotInstallationId(context) {
  const teamId = context.activity?.channelData?.team?.id;
  if (teamId) {
    return teamId;
  }

  // Fallback to use conversation id.
  // The conversation id is equal to team id only when the bot app is installed into the General channel.
  if (context.activity.conversation?.name === undefined && context.activity.conversation?.id) {
    return context.activity.conversation.id;
  }

  return undefined;
}

/**
 * A {@link NotificationTarget} that represents a team channel.
 *
 * @remarks
 * It's recommended to get channels from {@link TeamsBotInstallation.channels()}.
 */
class Channel {
  /**
   * Constructor.
   *
   * @remarks
   * It's recommended to get channels from {@link TeamsBotInstallation.channels()}, instead of using this constructor.
   *
   * @param parent - The parent {@link TeamsBotInstallation} where this channel is created from.
   * @param info - Detailed channel information.
   */
  constructor(parent, info) {
    this.parent = parent;
    this.info = info;
  }

  /**
   * Send a plain text message.
   *
   * @param text - The plain text message.
   * @param onError - An optional error handler that can catch exceptions during message sending.
   * If not defined, error will be handled by `BotAdapter.onTurnError`.
   *
   * @returns The response of sending message.
   */
  async sendMessage(text, onError) {
    const response = {};
    await this.parent.adapter.continueConversation(
      this.parent.conversationReference,
      async (context) => {
        const conversation = await this.newConversation(context);
        await this.parent.adapter.continueConversation(conversation, async (ctx) => {
          try {
            const res = await ctx.sendActivity(text);
            response.id = res?.id;
          } catch (error) {
            if (onError) {
              await onError(ctx, error);
            } else {
              throw error;
            }
          }
        });
      }
    );
    return response;
  }

  /**
   * Send an adaptive card message.
   *
   * @param card - The adaptive card raw JSON.
   * @param onError - An optional error handler that can catch exceptions during adaptive card sending.
   * If not defined, error will be handled by `BotAdapter.onTurnError`.
   *
   * @returns The response of sending adaptive card message.
   */
  async sendAdaptiveCard(card, onError) {
    const response = {};
    await this.parent.adapter.continueConversation(
      this.parent.conversationReference,
      async (context) => {
        const conversation = await this.newConversation(context);
        await this.parent.adapter.continueConversation(conversation, async (ctx) => {
          try {
            const res = await ctx.sendActivity({
              attachments: [CardFactory.adaptiveCard(card)],
              type: ActivityTypes.Message,
            });
            response.id = res?.id;
          } catch (error) {
            if (onError) {
              await onError(ctx, error);
            } else {
              throw error;
            }
          }
        });
      },
      true
    );
    return response;
  }

  /**
   * @internal
   */
  async newConversation(context) {
    const reference = context.activity.getConversationReference();
    const channelConversation = JSON.parse(JSON.stringify(reference));
    channelConversation.conversation.id = this.info.id || "";

    return Promise.resolve(channelConversation);
  }
}

/**
 * A {@link NotificationTarget} that represents a team member.
 *
 * @remarks
 * It's recommended to get members from {@link TeamsBotInstallation.members()}.
 */
class Member {
  /**
   * Constructor.
   *
   * @remarks
   * It's recommended to get members from {@link TeamsBotInstallation.members()}, instead of using this constructor.
   *
   * @param parent - The parent {@link TeamsBotInstallation} where this member is created from.
   * @param account - Detailed member account information.
   */
  constructor(parent, account) {
    this.parent = parent;
    this.account = account;
    /**
     * Notification target type. For member it's always "Person".
     */
    this.type = "personal";
  }

  /**
   * Send a plain text message.
   *
   * @param text - The plain text message.
   * @param onError - An optional error handler that can catch exceptions during message sending.
   * If not defined, error will be handled by `BotAdapter.onTurnError`.
   *
   * @returns The response of sending message.
   */
  async sendMessage(text, onError) {
    const response = {};
    await this.parent.adapter.continueConversation(
      this.parent.conversationReference,
      async (context) => {
        const conversation = await this.newConversation(context);
        await this.parent.adapter.continueConversation(conversation, async (ctx) => {
          try {
            const res = await ctx.sendActivity(text);
            response.id = res?.id;
          } catch (error) {
            if (onError) {
              await onError(ctx, error);
            } else {
              throw error;
            }
          }
        });
      }
    );
    return response;
  }

  /**
   * Send an adaptive card message.
   *
   * @param card - The adaptive card raw JSON.
   * @param onError - An optional error handler that can catch exceptions during adaptive card sending.
   * If not defined, error will be handled by `BotAdapter.onTurnError`.
   *
   * @returns The response of sending adaptive card message.
   */
  async sendAdaptiveCard(card, onError) {
    const response = {};
    await this.parent.adapter.continueConversation(
      this.parent.conversationReference,
      async (context) => {
        const conversation = await this.newConversation(context);
        await this.parent.adapter.continueConversation(conversation, async (ctx) => {
          try {
            const res = await ctx.sendActivity({
              attachments: [CardFactory.adaptiveCard(card)],
              type: ActivityTypes.Message,
            });
            response.id = res?.id;
          } catch (error) {
            if (onError) {
              await onError(ctx, error);
            } else {
              throw error;
            }
          }
        });
      },
      true
    );
    return response;
  }

  /**
   * @internal
   */ async newConversation(context) {
    const reference = context.activity.getConversationReference();
    const personalConversation = JSON.parse(JSON.stringify(reference));

    const connectorClient = context.turnState.get(this.parent.adapter.ConnectorClientKey);

    const conversationParams = {
      members: [this.account],
      isGroup: false,
      agent: context.activity.recipient,
      tenantId: context.activity.conversation.tenantId,
      activity: context.activity,
      channelData: context.activity.channelData,
    };
    const conversation = await connectorClient.createConversationAsync(conversationParams);
    personalConversation.conversation.id = conversation.id;

    return personalConversation;
  }
}

/**
 * A {@link NotificationTarget} that represents a bot installation. Teams Bot could be installed into
 * - Personal chat
 * - Group chat
 * - Team (by default the `General` channel)
 *
 * @remarks
 * It's recommended to get bot installations from {@link ConversationBot.installations()}.
 */
class TeamsBotInstallation {
  /**
   * Constructor
   *
   * @remarks
   * It's recommended to get bot installations from {@link ConversationBot.installations()}, instead of using this constructor.
   *
   * @param adapter - The bound `CloudAdapter`.
   * @param conversationReference - The bound `ConversationReference`.
   * @param botAppId - The bot app id.
   */
  constructor(adapter, conversationReference, botAppId) {
    /**
     * The bound `CloudAdapter`.
     */
    this.adapter = adapter;

    /**
     * The bound `ConversationReference`.
     */
    this.conversationReference = conversationReference;

    /**
     * Notification target type.
     *
     * @remarks
     * - "Channel" means bot is installed into a team and notification will be sent to its "General" channel.
     * - "Group" means bot is installed into a group chat.
     * - "Person" means bot is installed into a personal scope and notification will be sent to personal chat.
     */
    this.type = conversationReference.conversation.conversationType;

    /**
     * The bot app id.
     */
    this.botAppId = botAppId;
  }

  /**
   * Send a plain text message.
   *
   * @param text - The plain text message.
   * @param onError - An optional error handler that can catch exceptions during message sending.
   * If not defined, error will be handled by `BotAdapter.onTurnError`.
   *
   * @returns The response of sending message.
   */ async sendMessage(text, onError) {
    const response = {};
    await this.adapter.continueConversation(this.conversationReference, async (context) => {
      try {
        const res = await context.sendActivity(text);
        response.id = res?.id;
      } catch (error) {
        if (onError) {
          await onError(context, error);
        } else {
          throw error;
        }
      }
    });
    return response;
  }

  /**
   * Send an adaptive card message.
   *
   * @param card - The adaptive card raw JSON.
   * @param onError - An optional error handler that can catch exceptions during adaptive card sending.
   * If not defined, error will be handled by `BotAdapter.onTurnError`.
   *
   * @returns The response of sending adaptive card message.
   */ async sendAdaptiveCard(card, onError) {
    const response = {};
    await this.adapter.continueConversation(
      this.conversationReference,
      async (context) => {
        try {
          const adaptiveCard = {
            attachments: [CardFactory.adaptiveCard(card)],
            type: ActivityTypes.Message,
          };
          const res = await context.sendActivity(adaptiveCard);
          response.id = res?.id;
        } catch (error) {
          if (onError) {
            await onError(context, error);
          } else {
            throw error;
          }
        }
      },
      true
    );
    return response;
  }

  /**
   * Get channels from this bot installation.
   *
   * @returns An array of channels if bot is installed into a team, otherwise returns an empty array.
   */ async channels() {
    const channels = [];
    if (this.type !== "channel") {
      return channels;
    }

    let teamsChannels = [];
    await this.adapter.continueConversation(this.conversationReference, async (context) => {
      const teamId = getTeamsBotInstallationId(context);
      if (teamId !== undefined) {
        teamsChannels = await TeamsInfo.getTeamChannels(context, teamId);
      }
    });

    for (const channel of teamsChannels) {
      channels.push(new Channel(this, channel));
    }

    return channels;
  }

  /**
   * Gets a pagined list of members from this bot installation.
   *
   * @param pageSize - Suggested number of entries on a page.
   * @param continuationToken - A continuation token.
   * @returns An array of members from where the bot is installed.
   */ async getPagedMembers(pageSize, continuationToken) {
    let result = {
      data: [],
      continuationToken: "",
    };

    await this.adapter.continueConversation(this.conversationReference, async (context) => {
      const pagedMembers = await TeamsInfo.getPagedMembers(context, pageSize, continuationToken);

      result = {
        data: pagedMembers.members.map((m) => new Member(this, m)),
        continuationToken: pagedMembers.continuationToken,
      };
    });

    return result;
  }

  /**
   * Get team details from this bot installation
   *
   * @returns The team details if bot is installed into a team, otherwise returns `undefined`.
   */ async getTeamDetails() {
    if (this.type !== "channel") {
      return undefined;
    }

    let teamDetails;
    await this.adapter.continueConversation(this.conversationReference, async (context) => {
      const teamId = getTeamsBotInstallationId(context);
      if (teamId !== undefined) {
        teamDetails = await TeamsInfo.getTeamDetails(context, teamId);
      }
    });

    return teamDetails;
  }
}

/**
 * Provide utilities to send notification to varies targets (e.g., member, group, channel).
 */
class NotificationBot {
  /**
   * Constructor of the notification bot.
   *
   * @remarks
   * To ensure accuracy, it's recommended to initialize before handling any message.
   *
   * @param adapter - The bound `CloudAdapter`
   * @param options - The initialize options
   */
  constructor(adapter, storage, botAppId) {
    this.conversationReferenceStore = storage;
    this.adapter = adapter.use(new NotificationMiddleware(this.conversationReferenceStore));
    this.botAppId = botAppId;
  }

  /**
   * Create a {@link TeamsBotInstallation} instance with conversation reference.
   *
   * @param conversationReference - The bound `ConversationReference`.
   * @returns - The {@link TeamsBotInstallation} instance or null.
   */
  buildTeamsBotInstallation(conversationReference) {
    if (!conversationReference) {
      throw new Error("conversationReference is required.");
    }

    return new TeamsBotInstallation(this.adapter, conversationReference, this.botAppId);
  }
  /**
   * Validate the installation by getting paged memebers.
   *
   * @param conversationReference The bound `ConversationReference`.
   * @returns Returns false if recieves `BotNotInConversationRoster` error, otherwise returns true.
   */
  async validateInstallation(conversationReference) {
    let isValid = true;
    await this.adapter.continueConversation(conversationReference, async (context) => {
      try {
        // try get member to see if the installation is still valid
        await TeamsInfo.getPagedMembers(context, 1);
      } catch (error) {
        if (error.code === "BotNotInConversationRoster") {
          isValid = false;
        }
      }
    });
    return isValid;
  }

  /**
   * Gets a pagined list of targets where the bot is installed.
   *
   * @remarks
   * The result is retrieving from the persisted storage.
   *
   * @param pageSize - Suggested number of entries on a page.
   * @param continuationToken - A continuation token.
   *
   * @returns An array of {@link TeamsBotInstallation} with paged data and continuation token.
   */ async getPagedInstallations(pageSize, continuationToken, validationEnabled = true) {
    if (this.conversationReferenceStore === undefined || this.adapter === undefined) {
      throw new Error("NotificationBot has not been initialized.");
    }
    const references = await this.conversationReferenceStore.list(pageSize, continuationToken);
    const targets = [];
    for (const reference of references.data) {
      // validate connection
      let valid;
      if (validationEnabled) {
        // try get member to see if the installation is still valid
        valid = await this.validateInstallation(reference);
      }

      if (!validationEnabled || valid) {
        targets.push(new TeamsBotInstallation(this.adapter, reference, this.botAppId));
      } else {
        await this.conversationReferenceStore.delete([getKey(reference)]);
      }
    }

    return {
      data: targets,
      continuationToken: references.continuationToken,
    };
  }

  /**
   * Return the first {@link Member} where predicate is true, and undefined otherwise.
   *
   * @param predicate - Find calls predicate once for each member of the installation,
   * until it finds one where predicate returns true. If such a member is found, find
   * immediately returns that member. Otherwise, find returns undefined.
   * @param scope - The scope to find members from the installations
   * (personal chat, group chat, Teams channel).
   *
   * @returns The first {@link Member} where predicate is true, and `undefined` otherwise.
   */ async findMember(predicate, scope) {
    for (const target of await this.installations()) {
      if (this.matchSearchScope(target, scope)) {
        const members = [];
        let continuationToken;
        do {
          const pagedData = await target.getPagedMembers(undefined, continuationToken);
          continuationToken = pagedData.continuationToken;
          members.push(...pagedData.data);
        } while (continuationToken);

        for (const member of members) {
          if (await predicate(member)) {
            return member;
          }
        }
      }
    }

    return;
  }

  /**
   * Return the first {@link Channel} where predicate is true, and undefined otherwise.
   * (Ensure the bot app is installed into the `General` channel, otherwise undefined will be returned.)
   *
   * @param predicate - Find calls predicate once for each channel of the installation,
   * until it finds one where predicate returns true. If such a channel is found, find
   * immediately returns that channel. Otherwise, find returns `undefined`.
   *
   * @returns The first {@link Channel} where predicate is true, and `undefined` otherwise.
   */ async findChannel(predicate) {
    for (const target of await this.installations()) {
      if (target.type === "channel") {
        const teamDetails = await target.getTeamDetails();
        for (const channel of await target.channels()) {
          if (await predicate(channel, teamDetails)) {
            return channel;
          }
        }
      }
    }

    return;
  }

  /**
   * Return all {@link Member} where predicate is true, and empty array otherwise.
   *
   * @param predicate - Find calls predicate for each member of the installation.
   * @param scope - The scope to find members from the installations
   * (personal chat, group chat, Teams channel).
   *
   * @returns An array of {@link Member} where predicate is true, and empty array otherwise.
   */ async findAllMembers(predicate, scope) {
    const members = [];
    for (const target of await this.installations()) {
      if (this.matchSearchScope(target, scope)) {
        const targetMembers = [];
        let continuationToken;
        do {
          const pagedData = await target.getPagedMembers(undefined, continuationToken);
          continuationToken = pagedData.continuationToken;
          targetMembers.push(...pagedData.data);
        } while (continuationToken);

        for (const member of targetMembers) {
          if (await predicate(member)) {
            members.push(member);
          }
        }
      }
    }

    return members;
  }

  /**
   * Return all {@link Channel} where predicate is true, and empty array otherwise.
   * (Ensure the bot app is installed into the `General` channel, otherwise empty array will be returned.)
   *
   * @param predicate - Find calls predicate for each channel of the installation.
   *
   * @returns An array of {@link Channel} where predicate is true, and empty array otherwise.
   */ async findAllChannels(predicate) {
    const channels = [];
    for (const target of await this.installations()) {
      if (target.type === "channel") {
        const teamDetails = await target.getTeamDetails();
        for (const channel of await target.channels()) {
          if (await predicate(channel, teamDetails)) {
            channels.push(channel);
          }
        }
      }
    }

    return channels;
  }
  matchSearchScope(target, scope) {
    scope = scope ?? SearchScope.All;

    return (
      (target.type === "channel" && (scope & SearchScope.Channel) !== 0) ||
      (target.type === "groupChat" && (scope & SearchScope.Group) !== 0) ||
      (target.type === "personal" && (scope & SearchScope.Person) !== 0)
    );
  }

  /**
   * @internal
   * Get all targets where the bot is installed.
   *
   * @remarks
   * The result is retrieving from the persisted storage.
   *
   * @returns An array of {@link TeamsBotInstallation}
   */ async installations() {
    let continuationToken;
    const targets = [];
    do {
      const result = await this.getPagedInstallations(undefined, continuationToken);
      continuationToken = result.continuationToken;
      targets.push(...result.data);
    } while (continuationToken);

    return targets;
  }
}

/**
 * The search scope when calling {@link NotificationBot.findMember} and {@link NotificationBot.findAllMembers}.
 * The search scope is a flagged enum and it can be combined with `|`.
 * For example, to search from personal chat and group chat, use `SearchScope.Person | SearchScope.Group`.
 */
/**
 * The search scope when calling {@link NotificationBot.findMember} and {@link NotificationBot.findAllMembers}.
 * The search scope is a flagged enum and it can be combined with `|`.
 * For example, to search from personal chat and group chat, use `SearchScope.Person | SearchScope.Group`.
 */
const SearchScope = {
  /**
   * Search members from the installations in personal chat only.
   */
  Person: 1,

  /**
   * Search members from the installations in group chat only.
   */
  Group: 2,

  /**
   * Search members from the installations in Teams channel only.
   */
  Channel: 4,

  /**
   * Search members from all installations including personal chat, group chat and Teams channel.
   */
  All: 1 | 2 | 4,
};

module.exports = {
  getTeamsBotInstallationId,
  Channel,
  Member,
  TeamsBotInstallation,
  NotificationBot,
  SearchScope,
};
