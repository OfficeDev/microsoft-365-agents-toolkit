// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { ConversationReference, TurnContext } from "@microsoft/agents-hosting";
import { NotificationTargetType } from "./interface";

/**
 * @internal
 */
export function cloneConversation(
  conversation: Partial<ConversationReference>
): ConversationReference {
  return JSON.parse(JSON.stringify(conversation));
}

/**
 * @internal
 */
export function getKey(reference: Partial<ConversationReference>): string {
  // eslint-disable-next-line @typescript-eslint/restrict-template-expressions
  return `_${reference.conversation?.tenantId}_${reference.conversation?.id}`;
}

/**
 * @internal
 */
export function getTargetType(
  conversationReference: Partial<ConversationReference>
): NotificationTargetType | undefined {
  const conversationType = conversationReference.conversation?.conversationType;
  if (conversationType === "personal") {
    return NotificationTargetType.Person;
  } else if (conversationType === "groupChat") {
    return NotificationTargetType.Group;
  } else if (conversationType === "channel") {
    return NotificationTargetType.Channel;
  } else {
    return undefined;
  }
}

/**
 * @internal
 */
export function getTeamsBotInstallationId(context: TurnContext): string | undefined {
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
