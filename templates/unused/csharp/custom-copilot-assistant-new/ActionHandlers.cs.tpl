using Microsoft.Bot.Builder;
using Microsoft.Teams.AI.AI;
using Microsoft.Teams.AI.AI.Action;
using {{SafeProjectName}}.Model;

namespace {{SafeProjectName}}
{
    public class ActionHandlers
    {
        [Action(AIConstants.HttpErrorActionName)]
        public async Task<string> OnHttpError([ActionTurnContext] ITurnContext turnContext)
        {
            await turnContext.SendActivityAsync("An AI request failed. Please try again later.");
            return AIConstants.StopCommand;
        }

        [Action("createTask")]
        public async Task<string> OnCreateTask([ActionTurnContext] ITurnContext turnContext, [ActionTurnState] AppState state, [ActionParameters] Dictionary<string, object> entities)
        {
            string title = entities["title"].ToString();
            string description = entities["description"].ToString();
            MyTask task = new MyTask
            {
                Title = title,
                Description = description
            };
            Dictionary<string, MyTask> tasks = state.Conversation.Tasks;
            tasks[title] = task;
            state.Conversation.Tasks = tasks;
            return "task created, think about your next action";
        }

        [Action("deleteTask")]
        public async Task<string> OnDeleteTask([ActionTurnContext] ITurnContext turnContext, [ActionTurnState] AppState state, [ActionParameters] Dictionary<string, object> entities)
        {
            string title = entities["title"].ToString();
            Dictionary<string, MyTask> tasks = state.Conversation.Tasks;
            if (tasks.ContainsKey(title))
            {
                tasks.Remove(title);
                state.Conversation.Tasks = tasks;
                return "task has been deleted. Think about your next action";
            }
            else
            {
                await turnContext.SendActivityAsync($"There is no task '{title}'.");
                return "task not found, think about your next action";
            }
        }
    }
}
