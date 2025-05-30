import { TurnContext } from "botbuilder";
import { TurnState } from "@microsoft/teams-ai";

export async function resetMessage(context: TurnContext, state: TurnState) {
  state.deleteConversationState();
  await context.sendActivity("Ok lets start this over.");
}
