// In-memory storage for tasks - in production, use a persistent store
const taskStorage = new Map();

// Function to get or create conversation tasks
const getConversationTasks = (conversationId) => {
  if (!taskStorage.has(conversationId)) {
    taskStorage.set(conversationId, {});
  }
  return taskStorage.get(conversationId);
};

// Function to update conversation tasks
const updateConversationTasks = (conversationId, tasks) => {
  taskStorage.set(conversationId, tasks);
};

// Task handler functions
const createTaskHandler = async (parameters, conversationId) => {
  const { title, description } = parameters;
  console.log(`Creating task: ${title} for conversation ${conversationId}`);

  const tasks = getConversationTasks(conversationId);

  // Create new task
  const newTask = { title, description };
  tasks[title] = newTask;
  updateConversationTasks(conversationId, tasks);

  console.log(`Task "${title}" created successfully for conversation ${conversationId}`);
  return `Task "${title}" has been created successfully with description: "${description}"`;
};

const deleteTaskHandler = async (parameters, conversationId) => {
  const { title } = parameters;
  console.log(`Deleting task: ${title} for conversation ${conversationId}`);

  const tasks = getConversationTasks(conversationId);

  // Check if task exists
  if (tasks[title]) {
    delete tasks[title];
    updateConversationTasks(conversationId, tasks);

    console.log(`Task "${title}" deleted successfully for conversation ${conversationId}`);
    return `Task "${title}" has been deleted successfully.`;
  } else {
    return `Task "${title}" not found.`;
  }
};

const listTasksHandler = async (conversationId) => {
  console.log(`Listing all tasks for conversation ${conversationId}`);

  const tasks = getConversationTasks(conversationId);
  const taskList = Object.values(tasks);

  if (taskList.length === 0) {
    return "No tasks found.";
  }

  let response = `Current Tasks (${taskList.length}):\n\n`;
  taskList.forEach((task, index) => {
    response += `${task.title}\n   Description: ${task.description}\n\n`;
  });

  return response;
};

module.exports = {
  createTaskHandler,
  deleteTaskHandler,
  listTasksHandler,
  taskStorage,
};
