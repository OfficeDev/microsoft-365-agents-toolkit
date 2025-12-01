// Task interface for type safety
interface Task {
  title: string;
  description: string;
}

// In-memory storage for tasks - in production, use a persistent store
const taskStorage = new Map<string, Record<string, Task>>();

// Function to get or create conversation tasks
const getConversationTasks = (conversationId: string): Record<string, Task> => {
  if (!taskStorage.has(conversationId)) {
    taskStorage.set(conversationId, {});
  }
  return taskStorage.get(conversationId)!;
};

// Function to update conversation tasks
const updateConversationTasks = (conversationId: string, tasks: Record<string, Task>): void => {
  taskStorage.set(conversationId, tasks);
};

// Task handler functions
export const createTaskHandler = async (
  parameters: { title: string; description: string },
  conversationId: string
) => {
  const { title, description } = parameters;
  console.log(`Creating task: ${title} for conversation ${conversationId}`);

  const tasks = getConversationTasks(conversationId);

  // Create new task
  const newTask: Task = { title, description };
  tasks[title] = newTask;
  updateConversationTasks(conversationId, tasks);

  console.log(`Task "${title}" created successfully for conversation ${conversationId}`);
  return `Task "${title}" has been created successfully with description: "${description}"`;
};

export const deleteTaskHandler = async (parameters: { title: string }, conversationId: string) => {
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

export const listTasksHandler = async (conversationId: string) => {
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

// Export task storage utilities for reset functionality
export { taskStorage };
