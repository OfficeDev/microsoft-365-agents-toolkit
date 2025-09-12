import { Activity, ActivityTypes, TurnContext } from '@microsoft/teams.common';
import { AppInterface, ConfigInterface, createApp } from '@microsoft/teams.apps';

interface ConversationState {
  count: number;
}

// Create application using declarative configuration
export async function createTeamsApp(): Promise<AppInterface> {
  const app = await createApp<ConversationState>({
    storage: {
      type: 'memory'
    }
  });

  // Listen for user to say '/reset' and then delete conversation state
  app.onMessage('/reset', async (context: TurnContext, state: ConversationState) => {
    state.count = 0;
    await context.sendActivity("Ok I've reset the conversation state.");
  });

  app.onMessage('/count', async (context: TurnContext, state: ConversationState) => {
    const count = state.count ?? 0;
    await context.sendActivity(`The count is ${count}`);
  });

  app.onMessage('/diag', async (context: TurnContext) => {
    await context.sendActivity(JSON.stringify(context.activity));
  });

  app.onMessage('/state', async (context: TurnContext, state: ConversationState) => {
    await context.sendActivity(JSON.stringify(state));
  });

  app.onMessage('/runtime', async (context: TurnContext) => {
    const runtime = {
      nodeversion: process.version,
    };
    await context.sendActivity(JSON.stringify(runtime));
  });

  app.onConversationUpdate(
    'membersAdded',
    async (context: TurnContext) => {
      await context.sendActivity(
        "Hi there! I'm an echo bot running on Teams AI Library V2 that will echo what you said to me."
      );
    }
  );

  // Listen for ANY message to be received. MUST BE AFTER ANY OTHER MESSAGE HANDLERS
  app.onActivity(
    ActivityTypes.Message,
    async (context: TurnContext, state: ConversationState) => {
      // Increment count state
      let count = state.count ?? 0;
      state.count = ++count;

      // Echo back users request
      await context.sendActivity(`[${count}] you said: ${context.activity.text}`);
    }
  );

  return app;
}