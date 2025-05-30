import { TurnContext } from "botbuilder";
import { ApplicationTurnState } from "./turnState";

export async function resetMessage(
  context: TurnContext,
  state: ApplicationTurnState
): Promise<void> {
  state.deleteConversationState();
  await context.sendActivity("Ok lets start this over.");
}
