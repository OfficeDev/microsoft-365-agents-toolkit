# Automated Testing with playground-cli

Run programmatic smoke tests and integration tests against your bot using `@microsoft/m365agentsplayground-cli` — a headless, browser-free API that drives bot conversations via code. Use this for CI pipelines, multi-turn conversation testing, and verifying bot behavior without manual interaction.

## Setup

The `@microsoft/m365agentsplayground-cli` package lives in the [microsoft-365-agents-playground](https://github.com/OfficeDev/microsoft-365-agents-playground) monorepo.

```bash
# 1. Build the package from the playground monorepo
cd /path/to/microsoft-365-agents-playground
npm run build --workspace=packages/playground-cli

# 2. In your bot project, add as a dev dependency using a file: path
npm install --save-dev file:/path/to/microsoft-365-agents-playground/packages/playground-cli
```

> **Alternative (ConversationServer):** For ad-hoc smoke tests without modifying your project's `package.json`, start the ConversationServer directly from the playground monorepo and call it via HTTP — no install needed in the bot project:
> ```bash
> cd /path/to/microsoft-365-agents-playground
> npm run server --workspace=packages/playground-cli -- --port 9000
> ```

## Smoke Test Workflow

> This is the workflow used in the **Combined Workflow** in `test-playground.md`. Follow these steps to generate and run smoke tests for any bot.

### Step 1: Inspect the bot

Read the bot's entry point (`app.ts`, `index.ts`, or similar) to understand:
- What message handlers are registered (`app.message`, `app.command`, `app.activity`)
- Whether it uses `ChatPrompt` / AI (read the system prompt to understand its purpose)
- Any known commands or keywords it responds to

### Step 2: Design test cases

Based on what you found, design 5–7 smoke tests covering:

| Priority | Test | Typical Input |
|---|---|---|
| P0 | Greeting / welcome | `"Hello"` |
| P0 | Main feature | Input that exercises the bot's primary purpose |
| P1 | Known command | `/help`, `/start`, or first listed command |
| P1 | Follow-up / context | Second message in same conversation |
| P2 | Edge case / unknown input | `"xyzzy"`, `"???"` |
| P2 | Bot install event | `turn_type: "install"` (if bot has install handler) |

### Step 3: Create and run a smoke test script

Write a temporary script `_smoke-test.mjs` in the project root:

```js
import { TestClient } from "@microsoft/m365agentsplayground-cli";

const BOT = "http://localhost:3978/api/messages";

const tests = [
  { id: "T1", name: "Greeting",       input: "Hello" },
  { id: "T2", name: "Main feature",   input: "What can you do?" },
  { id: "T3", name: "Help command",   input: "/help" },
  { id: "T4", name: "Follow-up",      input: "Tell me more" },
  { id: "T5", name: "Unknown input",  input: "xyzzy" },
];

const client = new TestClient({ botEndpoint: BOT, timeout: 20000, deliveryMode: "expectReplies" });
await client.start();

const results = [];
for (const t of tests) {
  client.newConversation();
  try {
    const replies = await client.sendMessage(t.input);
    const text = replies[0]?.text ?? (replies[0]?.attachments?.length ? "[Card]" : "[empty]");
    results.push({ ...t, status: "PASS", response: text.slice(0, 100) });
  } catch (e) {
    results.push({ ...t, status: "FAIL", response: e.message.slice(0, 100) });
  }
}

await client.stop();

// Print as markdown table
console.log("| # | Test | Input | Status | Response |");
console.log("|---|------|-------|--------|---------|");
for (const r of results) {
  const icon = r.status === "PASS" ? "✅" : "❌";
  console.log(`| ${r.id} | ${r.name} | \`${r.input}\` | ${icon} ${r.status} | ${r.response} |`);
}
```

```bash
node _smoke-test.mjs
```

### Step 4: Present results as a table

Parse the output and present it to the user as a markdown table, for example:

| # | Test | Input | Status | Response |
|---|------|-------|--------|---------|
| T1 | Greeting | `Hello` | ✅ PASS | Hello! I'm your assistant... |
| T2 | Main feature | `What can you do?` | ✅ PASS | I can help you with... |
| T3 | Help command | `/help` | ✅ PASS | [Card] |
| T4 | Follow-up | `Tell me more` | ✅ PASS | Sure! Here are the details... |
| T5 | Unknown input | `xyzzy` | ✅ PASS | I didn't understand that. Try... |

Clean up when done:
```bash
rm _smoke-test.mjs
# or on Windows:
del _smoke-test.mjs
```

---

## API Reference

### TestClient

```typescript
import { TestClient } from "@microsoft/m365agentsplayground-cli";

const client = new TestClient({
  botEndpoint: "http://localhost:3978/api/messages",
  timeout: 15000,          // ms to wait for bot reply (default: 5000)
  deliveryMode: "expectReplies",  // required for teams-ai / teams.ts bots
  streamingSettleDelayMs: 800,    // ms quiet-period for streaming bots
});

await client.start();            // must call before sendMessage()
await client.stop();             // call in teardown / after all tests
client.newConversation();        // reset conversationId between test cases
const id = client.getConversationId();

const replies: BotResponse[] = await client.sendMessage("Hello");
// replies is an array — bots can send multiple activities per turn
const text = replies[0]?.text;
const card = replies[0]?.attachments?.[0];
const allMessages = client.getMessages();
const last = client.getLastBotMessage();
```

### BotResponse fields

| Field | Type | Description |
|---|---|---|
| `text` | `string` | Plain text content |
| `attachments` | `Attachment[]` | Adaptive Cards, Hero Cards, etc. |
| `suggestedActions` | `object` | Suggested action buttons |
| `type` | `string` | Activity type (`message`, `typing`, etc.) |

### ConversationServer (HTTP API)

For non-Node test runners (Python, curl, any language):

```bash
# Start the server (background process)
cd /path/to/microsoft-365-agents-playground
npm run server --workspace=packages/playground-cli -- --port 9000

# Or programmatically (Node):
# import { createConversationServer } from "@microsoft/m365agentsplayground-cli";
# const server = createConversationServer({ port: 9000 });

# Verify health
curl http://localhost:9000/health
# → {"status":"ok"}
```

**POST /run-conversation**

```json
{
  "config": {
    "botEndpoint": "http://localhost:3978/api/messages",
    "timeout": 30000,
    "deliveryMode": "expectReplies",
    "personas": {
      "alice": { "id": "user-alice", "name": "Alice", "email": "alice@example.com" }
    }
  },
  "scenario": "smoke-test",
  "input": {
    "turns": [
      { "test_id": "t1", "prompt": "Hello" },
      { "test_id": "t2", "prompt": "What can you do?", "turn_type": "chat" },
      { "test_id": "t3", "prompt": "", "turn_type": "install" },
      { "test_id": "t4", "prompt": "<html>Order shipped</html>", "turn_type": "sendEmail", "persona": "alice" }
    ]
  }
}
```

**Response:**

```json
{
  "turns": [
    { "test_id": "t1", "status": "Completed", "actual_response": "Hello! I'm your assistant..." },
    { "test_id": "t2", "status": "Completed", "actual_response": "I can help you with..." },
    { "test_id": "t3", "status": "Completed", "actual_response": "Welcome! I'm installed..." },
    { "test_id": "t4", "status": "TimedOut",   "actual_response": "" }
  ]
}
```

Turn statuses: `Completed` | `TimedOut` | `Errored` | `Skipped` (skipped = previous turn failed)

## Turn Types

| `turn_type` | Simulates |
|---|---|
| `"chat"` | Normal user message (default) |
| `"sendEmail"` | Email notification received |
| `"mentionInWord"` | @mention in Word document |
| `"install"` | Bot installation event |
| `"userAdded"` | Member added to conversation |
| `"botAdded"` | Bot added to team |
| `"channelCreated"` | New channel created |
| `"teamRenamed"` | Team renamed |

## Key Configuration Notes

| Setting | When to use |
|---|---|
| `deliveryMode: "expectReplies"` | Required for `@microsoft/teams-ai` and `teams.ts` bots |
| `timeout: 30000` | Increase for bots calling external APIs or LLMs |
| `streamingSettleDelayMs: 2000` | Increase only if LLM pauses > 800ms between stream chunks |
| `personas` | Required for testing notification bots with specific `from` identity |

## Common Pitfalls

- **No `await client.start()`** → `sendMessage()` throws immediately
- **`deliveryMode` missing for teams-ai bots** → `sendMessage()` returns `[]`
- **No `client.newConversation()` between tests** → bot state bleeds between test cases
- **ConversationServer not started before HTTP tests** → connection refused; verify `/health` first
- **Parallel tests sharing one `TestClient`** → mixed-up responses; use separate instances

## References

- [playground-cli README](https://github.com/OfficeDev/microsoft-365-agents-playground/blob/main/packages/playground-cli/README.md)
- [playground-cli-example sample](https://github.com/OfficeDev/microsoft-365-agents-playground/tree/main/samples/playground-cli-example)
- Manual interactive testing → [playground.md](playground.md)
