// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

const ActivityType = {
  CurrentBotInstalled: 0,
  CurrentBotMessaged: 1,
  CurrentBotUninstalled: 2,
  TeamDeleted: 3,
  TeamRestored: 4,
  Unknown: 5,
};

function getKey(reference) {
  return `_${reference.conversation?.tenantId}_${reference.conversation?.id}`;
}

class NotificationMiddleware {
  constructor(storage) {
    this.conversationReferenceStore = storage;
  }

  async onTurn(context, next) {
    const type = this.classifyActivity(context.activity);
    switch (type) {
      case ActivityType.CurrentBotInstalled:
      case ActivityType.TeamRestored: {
        const reference = context.activity.getConversationReference();
        await this.conversationReferenceStore.write({
          [getKey(reference)]: reference,
        });
        break;
      }
      case ActivityType.CurrentBotMessaged: {
        await this.tryAddMessagedReference(context);
        break;
      }
      case ActivityType.CurrentBotUninstalled:
      case ActivityType.TeamDeleted: {
        const reference = context.activity.getConversationReference();
        await this.conversationReferenceStore.delete([getKey(reference)]);
        break;
      }
      default:
        break;
    }

    await next();
  }

  classifyActivity(activity) {
    const activityType = activity.type;
    if (activityType === "installationUpdate") {
      const action = activity.action?.toLowerCase();
      if (action === "add" || action === "add-upgrade") {
        return ActivityType.CurrentBotInstalled;
      } else {
        return ActivityType.CurrentBotUninstalled;
      }
    } else if (activityType === "conversationUpdate") {
      const eventType = activity.channelData?.eventType;
      if (eventType === "teamDeleted") {
        return ActivityType.TeamDeleted;
      } else if (eventType === "teamRestored") {
        return ActivityType.TeamRestored;
      }
    } else if (activityType === "message") {
      return ActivityType.CurrentBotMessaged;
    }

    return ActivityType.Unknown;
  }

  async tryAddMessagedReference(context) {
    const reference = context.activity.getConversationReference();
    const conversationType = reference?.conversation?.conversationType;
    if (conversationType === "personal" || conversationType === "groupChat") {
      await this.conversationReferenceStore.write({ [getKey(reference)]: reference });
    } else if (conversationType === "channel") {
      const teamId = context.activity?.channelData?.team?.id;
      const channelId = context.activity.channelData?.channel?.id;
      // `teamId === channelId` means General channel. Ignore messaging in non-General channel.
      if (teamId !== undefined && (channelId === undefined || teamId === channelId)) {
        const teamReference = JSON.parse(JSON.stringify(reference));
        teamReference.conversation.id = teamId;
        await this.conversationReferenceStore.write({ [getKey(teamReference)]: teamReference });
      }
    }
  }
}

module.exports = {
  getKey,
  NotificationMiddleware,
};
