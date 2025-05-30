import { DefaultConversationState, TurnState } from "@microsoft/teams-ai";

export interface Task {
  title: string;
  description: string;
}

export interface ConversationState extends DefaultConversationState {
  tasks: Record<string, Task>;
}

export type ApplicationTurnState = TurnState<ConversationState>;
