const { ActivityTypes } = require("@microsoft/agents-activity");
const {
  AgentApplication,
  AttachmentDownloader,
  MemoryStorage,
} = require("@microsoft/agents-hosting");
const { version } = require("@microsoft/agents-hosting/package.json");

const downloader = new AttachmentDownloader();

// Define storage and application
const storage = new MemoryStorage();
const agentApp = new AgentApplication({
  storage,
  fileDownloaders: [downloader],
});

// Listen for user to say '/reset' and then delete conversation state
agentApp.onMessage("/reset", async (context, state) => {
  state.deleteConversationState();
  await context.sendActivity("Ok I've deleted the current conversation state.");
});

agentApp.onMessage("/count", async (context, state) => {
  const count = state.conversation.count ?? 0;
  await context.sendActivity(`The count is ${count}`);
});

agentApp.onMessage("/diag", async (context, state) => {
  await state.load(context, storage);
  await context.sendActivity(JSON.stringify(context.activity));
});

agentApp.onMessage("/state", async (context, state) => {
  await state.load(context, storage);
  await context.sendActivity(JSON.stringify(state));
});

agentApp.onMessage("/runtime", async (context, state) => {
  const runtime = {
    nodeversion: process.version,
    sdkversion: version,
  };
  await context.sendActivity(JSON.stringify(runtime));
});

agentApp.onConversationUpdate("membersAdded", async (context, state) => {
  await context.sendActivity(
    `Hi there! I'm an echo bot running on Agents SDK version ${version} that will echo what you said to me.`
  );
});

// Listen for ANY message to be received. MUST BE AFTER ANY OTHER MESSAGE HANDLERS
agentApp.onActivity(ActivityTypes.Message, async (context, state) => {
  // Increment count state
  let count = state.conversation.count ?? 0;
  state.conversation.count = ++count;

  // Echo back users request
  await context.sendActivity(`[${count}] you said: ${context.activity.text}`);
});

agentApp.onActivity(/^message/, async (context, state) => {
  await context.sendActivity(`Matched with regex: ${context.activity.type}`);
});

agentApp.onActivity(
  async (context) => Promise.resolve(context.activity.type === "message"),
  async (context, state) => {
    await context.sendActivity(`Matched function: ${context.activity.type}`);
  }
);

module.exports.agentApp = agentApp;
