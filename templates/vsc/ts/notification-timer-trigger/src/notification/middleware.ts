// Copyright (c) Microsoft Corporation.Add commentMore actions
// Licensed under the MIT license.

import { Activity, ConversationReference } from "@microsoft/agents-activity";
import { TurnContext, Middleware } from "@microsoft/agents-hosting";
import { IStorage } from "./interface";

/**
 * @internal
 */
enum ActivityType {
  CurrentBotInstalled,
  CurrentBotMessaged,
  CurrentBotUninstalled,
  TeamDeleted,
  TeamRestored,
  Unknown,
}

export function getKey(reference: Partial<ConversationReference>): string {
  // eslint-disable-next-line @typescript-eslint/restrict-template-expressions
  return `_${reference.conversation?.tenantId}_${reference.conversation?.id}`;
}

export class NotificationMiddleware implements Middleware {
  private readonly conversationReferenceStore: IStorage;

  constructor(storage: IStorage) {
    this.conversationReferenceStore = storage;
  }

  public async onTurn(context: TurnContext, next: () => Promise<void>): Promise<void> {
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

  private classifyActivity(activity: Activity): ActivityType {
    const activityType = activity.type;
    if (activityType === "installationUpdate") {
      const action = activity.action?.toLowerCase();
      if (action === "add" || action === "add-upgrade") {
        return ActivityType.CurrentBotInstalled;
      } else {
        return ActivityType.CurrentBotUninstalled;
      }
    } else if (activityType === "conversationUpdate") {
      const eventType = activity.channelData?.eventType as string;
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

  private async tryAddMessagedReference(context: TurnContext): Promise<void> {
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
