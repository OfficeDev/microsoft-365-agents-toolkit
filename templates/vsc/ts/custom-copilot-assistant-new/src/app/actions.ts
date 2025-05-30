import { TurnContext } from "botbuilder";
import { ApplicationTurnState, Task } from "./turnState";

interface CreateTaskParameters {
  title: string;
  description: string;
}

interface DeleteTaskParameters {
  title: string;
}

export async function createTask(
  context: TurnContext,
  state: ApplicationTurnState,
  parameters: CreateTaskParameters
): Promise<string> {
  if (state.conversation.tasks === undefined) {
    state.conversation.tasks = {};
  }
  const task: Task = {
    title: parameters.title,
    description: parameters.description,
  };
  state.conversation.tasks[parameters.title] = task;
  return "task created, think about your next action";
}

export async function deleteTask(
  context: TurnContext,
  state: ApplicationTurnState,
  parameters: DeleteTaskParameters
): Promise<string> {
  if (state.conversation.tasks === undefined) {
    state.conversation.tasks = {};
  }
  if (state.conversation.tasks[parameters.title] === undefined) {
    await context.sendActivity(`There is no task '${parameters.title}'.`);
    return "task not found, think about your next action";
  }
  delete state.conversation.tasks[parameters.title];
  return "task deleted, think about your next action";
}
