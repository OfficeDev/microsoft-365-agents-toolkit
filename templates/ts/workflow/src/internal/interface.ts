import { TurnState } from "@microsoft/agents-hosting";

export interface ConversationState {
  count: number;
}
export type ApplicationTurnState = TurnState<ConversationState>;
