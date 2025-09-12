const { ActivityTypes } = require('@microsoft/teams.common');
const { createApp } = require('@microsoft/teams.apps');

// Create application using declarative configuration
async function createTeamsApp() {
  const app = await createApp({
    storage: {
      type: 'memory'
    }
  });

  // Listen for user to say '/reset' and then delete conversation state
  app.onMessage('/reset', async (context, state) => {
    state.count = 0;
    await context.sendActivity("Ok I've reset the conversation state.");
  });

  app.onMessage('/count', async (context, state) => {
    const count = state.count ?? 0;
    await context.sendActivity(`The count is ${count}`);
  });

  app.onMessage('/diag', async (context) => {
    await context.sendActivity(JSON.stringify(context.activity));
  });

  app.onMessage('/state', async (context, state) => {
    await context.sendActivity(JSON.stringify(state));
  });

  app.onMessage('/runtime', async (context) => {
    const runtime = {
      nodeversion: process.version,
    };
    await context.sendActivity(JSON.stringify(runtime));
  });

  app.onConversationUpdate('membersAdded', async (context) => {
    await context.sendActivity(
      "Hi there! I'm an echo bot running on Teams AI Library V2 that will echo what you said to me."
    );
  });

  // Listen for ANY message to be received. MUST BE AFTER ANY OTHER MESSAGE HANDLERS
  app.onActivity(ActivityTypes.Message, async (context, state) => {
    // Increment count state
    let count = state.count ?? 0;
    state.count = ++count;

    // Echo back users request
    await context.sendActivity(`[${count}] you said: ${context.activity.text}`);
  });

  return app;
}

module.exports.createTeamsApp = createTeamsApp;