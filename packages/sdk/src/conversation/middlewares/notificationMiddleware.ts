// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { Activity } from "@microsoft/agents-activity";
import { TurnContext } from "@microsoft/agents-hosting";
import { Middleware } from "@microsoft/agents-hosting/dist/src/middlewareSet";
import { cloneConversation, getKey } from "../utils";
import { ConversationReferenceStore } from "../interface";

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

/**
 * @internal
 */
export interface NotificationMiddlewareOptions {
  conversationReferenceStore: ConversationReferenceStore;
}

/**
 * @internal
 */
export class NotificationMiddleware implements Middleware {
  private readonly conversationReferenceStore: ConversationReferenceStore;

  constructor(options: NotificationMiddlewareOptions) {
    this.conversationReferenceStore = options.conversationReferenceStore;
  }

  public async onTurn(context: TurnContext, next: () => Promise<void>): Promise<void> {
    const type = this.classifyActivity(context.activity);
    switch (type) {
      case ActivityType.CurrentBotInstalled:
      case ActivityType.TeamRestored: {
        const reference = context.activity.getConversationReference();
        await this.conversationReferenceStore.add(getKey(reference), reference, {
          overwrite: true,
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
        await this.conversationReferenceStore.remove(getKey(reference), reference);
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
      await this.conversationReferenceStore.add(getKey(reference), reference, { overwrite: false });
    } else if (conversationType === "channel") {
      const teamId = context.activity?.channelData?.team?.id;
      const channelId = context.activity.channelData?.channel?.id;
      // `teamId === channelId` means General channel. Ignore messaging in non-General channel.
      if (teamId !== undefined && (channelId === undefined || teamId === channelId)) {
        const teamReference = cloneConversation(reference);
        teamReference.conversation.id = teamId;
        await this.conversationReferenceStore.add(getKey(teamReference), teamReference, {
          overwrite: false,
        });
      }
    }
  }
}
