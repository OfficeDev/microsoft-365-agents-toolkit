async function createTask(context, state, parameters) {
  if (state.conversation.tasks === undefined) {
    state.conversation.tasks = {};
  }
  const task = {
    title: parameters.title,
    description: parameters.description,
  };
  state.conversation.tasks[parameters.title] = task;
  return "task created, think about your next action";
}

async function deleteTask(context, state, parameters) {
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

module.exports = {
  createTask,
  deleteTask,
};
