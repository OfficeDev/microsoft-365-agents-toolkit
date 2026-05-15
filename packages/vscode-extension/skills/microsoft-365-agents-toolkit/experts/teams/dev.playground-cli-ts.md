# dev.playground-cli-ts

## purpose

Programmatic integration testing of Teams bots/agents using `@microsoft/m365agentsplayground-cli` — a headless, browser-free library that drives bot conversations via code, suitable for automated tests and CI pipelines.

## rules

1. **Use `@microsoft/m365agentsplayground-cli` for automated testing.** This is the programmatic counterpart to the `agentsplayground` browser UI. Import `TestClient` from the package and call `sendMessage()` to drive conversations without a browser. It connects to the same `http://localhost:{PORT}/api/messages` endpoint. [github.com/OfficeDev/microsoft-365-agents-playground](https://github.com/OfficeDev/microsoft-365-agents-playground)

2. **Always call `client.start()` before sending messages and `client.stop()` in teardown.** `start()` initializes the WebSocket connection and conversation state. Without it, `sendMessage()` throws. Always clean up with `stop()` in `after()`/`afterAll()` to prevent hanging test processes.

3. **`sendMessage()` returns `BotResponse[]` — always destructure or iterate.** A single user message can produce multiple bot replies (e.g., a text message followed by a card). Destructure with `const [first, second] = await client.sendMessage("Hello")` or loop over the array. Never assume exactly one response.

4. **Set `timeout` in `TestClientConfig` for slow bots.** Default is 5000 ms. Increase for bots that call external APIs or LLMs: `new TestClient({ botEndpoint: "...", timeout: 30000 })`. A `TimedOut` status in results means the bot did not respond within the limit.

5. **Use `ConversationServer` (HTTP wrapper) for non-Node test runners.** Start the server with `npm run server --workspace=packages/playground-cli -- --port 9000` or `createConversationServer({ port: 9000 })`. Then call `POST /run-conversation` with a JSON payload from Python, curl, or any HTTP client. Check `GET /health` to verify the server is up.

6. **Use `turn_type` to simulate non-chat activities.** Default turn type is `"chat"`. Use `"sendEmail"` to test email notification handling, `"install"` for bot installation events, `"userAdded"` for member-added activities, `"channelCreated"` and `"teamRenamed"` for team-scoped events. Mix types in a single conversation by setting `turn_type` per turn.

7. **Use `personas` for notification turns that require a specific `from` identity.** Define named personas in `config.personas` with `{ id, name, email }`. Reference by name in `turn.persona`. The persona's `id` is used as `from.id` on the activity. Essential for testing bots that respond differently to different senders.

8. **Use `deliveryMode: "expectReplies"` for bots that return all replies in the HTTP response.** Teams AI library bots and teams.ts bots default to `expectReplies`. Set this when the bot delivers all replies inline rather than via async proactive sends. Without it, `sendMessage()` may miss replies.

9. **Streaming bots are handled automatically.** When the simulator receives `"Loading stream results..."` as the response text, it switches to streaming mode and waits for a `streamType:"final"` WebSocket event. If the bot doesn't emit `streamType`, it falls back to a quiet-period wait (`streamingSettleDelayMs`, default 800 ms). Increase `streamingSettleDelayMs` only for slow LLMs with long pauses between stream chunks.

10. **Multiple `TestClient` instances in the same process are safe.** Each instance gets a unique `conversationId`, so parallel tests don't interfere. To test multiple bot endpoints in parallel, run separate `ConversationServer` processes on different ports and point each test suite at its own server URL.

11. **Use `client.newConversation()` to reset conversation state between test cases.** This resets the `conversationId` without creating a new `TestClient`. Reuse the same client across tests within a suite, but call `newConversation()` between independent scenarios to avoid state leakage.

12. **Check `turn.status` in `ConversationServer` results before asserting on `actual_response`.** Possible statuses: `Completed`, `TimedOut`, `Errored`, `Skipped`. A `Skipped` turn means a previous turn `Errored` or `TimedOut`. Assert `status === "Completed"` first to get clear failure messages.

## patterns

### Basic TestClient integration test (Mocha/TypeScript)

```typescript
import { TestClient } from "@microsoft/m365agentsplayground-cli";
import { expect } from "chai";

describe("MyBot", () => {
  let client: TestClient;

  before(async () => {
    client = new TestClient({
      botEndpoint: "http://localhost:3978/api/messages",
      timeout: 15000,
      deliveryMode: "expectReplies",
    });
    await client.start();
  });

  after(async () => {
    await client.stop();
  });

  beforeEach(() => {
    client.newConversation(); // isolate each test
  });

  it("should greet the user", async () => {
    const [response] = await client.sendMessage("Hello");
    expect(response.text).to.include("Hello");
  });

  it("should return a card for /help", async () => {
    const responses = await client.sendMessage("/help");
    const cardResponse = responses.find((r) => r.attachments?.length);
    expect(cardResponse).to.exist;
    expect(cardResponse!.attachments![0].contentType).to.equal(
      "application/vnd.microsoft.card.adaptive"
    );
  });
});
```

### ConversationServer: multi-turn conversation via HTTP (Python example)

```python
import subprocess
import requests
import time

# Start the server (in CI, run this as a background process before tests)
# npm run server --workspace=packages/playground-cli -- --port 9000

# Verify the server is healthy
resp = requests.get("http://localhost:9000/health")
assert resp.json()["status"] == "ok"

# Run a multi-turn conversation
result = requests.post("http://localhost:9000/run-conversation", json={
    "config": {
        "botEndpoint": "http://localhost:3978/api/messages",
        "timeout": 30000,
        "deliveryMode": "expectReplies",
    },
    "scenario": "smoke-test",
    "input": {
        "turns": [
            {"test_id": "t1", "prompt": "Hello"},
            {"test_id": "t2", "prompt": "What can you do?"},
        ]
    },
}).json()

for turn in result["turns"]:
    assert turn["status"] == "Completed", f"Turn {turn['test_id']} failed: {turn['status']}"
    print(f"[{turn['test_id']}] {turn['actual_response'][:80]}")
```

### ConversationServer: testing notification activities (email/mention)

```json
{
  "config": {
    "botEndpoint": "http://localhost:3978/api/messages",
    "timeout": 30000,
    "personas": {
      "customer": { "id": "customer-id", "name": "Alice", "email": "alice@example.com" }
    }
  },
  "scenario": "email-notification-test",
  "input": {
    "turns": [
      {
        "test_id": "t1",
        "prompt": "<html><body>Order #123 has been shipped.</body></html>",
        "turn_type": "sendEmail",
        "persona": "customer"
      },
      {
        "test_id": "t2",
        "prompt": "Summarize the last email I received",
        "turn_type": "chat"
      }
    ]
  }
}
```

### ConversationServer: mixed chat types within one conversation

```json
{
  "config": {
    "botEndpoint": "http://localhost:3978/api/messages",
    "chatType": "personal"
  },
  "scenario": "channel-test",
  "input": {
    "turns": [
      { "test_id": "t1", "prompt": "Hello from personal chat", "turn_type": "chat" },
      { "test_id": "t2", "prompt": "@bot Hello from channel", "turn_type": "chat", "chat_type": "channel" }
    ]
  }
}
```

### Bot installation event test

```typescript
import { TestClient } from "@microsoft/m365agentsplayground-cli";

const client = new TestClient({ botEndpoint: "http://localhost:3978/api/messages" });
await client.start();

// Simulate bot installation
const responses = await client.sendMessage("", { turn_type: "install" } as any);
// Bot should send a welcome message on install
const welcome = responses.find((r) => r.text?.toLowerCase().includes("welcome"));
expect(welcome).to.exist;

await client.stop();
```

### Streaming bot test

```typescript
import { TestClient } from "@microsoft/m365agentsplayground-cli";

const client = new TestClient({
  botEndpoint: "http://localhost:3978/api/messages",
  timeout: 60000,
  // Increase only if your LLM pauses mid-stream for > 800ms
  streamingSettleDelayMs: 2000,
} as any);
await client.start();

const [response] = await client.sendMessage("Tell me a long story");
// For streaming bots, the full assembled text is in response.text
expect(response.text).to.have.length.greaterThan(0);
await client.stop();
```

## pitfalls

- **Forgetting `await client.start()`**: `sendMessage()` will throw if the client was never started. Always call `start()` in `before()`/`beforeAll()` and `stop()` in `after()`/`afterAll()`.
- **Assuming one reply per message**: Many bots send multiple activities per user turn (e.g., a typing indicator, then a text, then a card). `sendMessage()` returns an array — always check all items, not just the first.
- **Not isolating tests with `newConversation()`**: Without resetting conversation state between tests, bot context from a previous test contaminates the next. Call `client.newConversation()` in `beforeEach()`.
- **Wrong `deliveryMode`**: If using a teams.ts or teams-ai bot with `stream: false`, the default `deliveryMode` is usually `"expectReplies"`. Omitting this means replies are sent asynchronously and `sendMessage()` returns an empty array.
- **Not checking `turn.status` before asserting response content**: If a turn `TimedOut` or `Errored`, `actual_response` may be empty or undefined. Always assert `status === "Completed"` first for clear test failure messages.
- **ConversationServer not running when HTTP tests start**: The server must be started (and `/health` verified) before HTTP-based tests run. In CI, start the server as a background step before the test step.
- **Parallel tests sharing a `TestClient`**: A single `TestClient` has one `conversationId`. Parallel test cases using the same client will mix up responses. Use separate `TestClient` instances or `newConversation()` sequentially.
- **Building from source without installing first**: The package lives in the playground monorepo. Run `npm run build --workspace=packages/playground-cli` from the repo root before using it via `file:` dependency.

## references

- [playground-cli README](https://github.com/OfficeDev/microsoft-365-agents-playground/blob/main/packages/playground-cli/README.md)
- [playground-cli-example](https://github.com/OfficeDev/microsoft-365-agents-playground/tree/main/samples/playground-cli-example)
- [Agents Playground browser UI](https://github.com/OfficeDev/microsoft-365-agents-toolkit/releases) — the browser-based counterpart; use `agentsplayground` CLI for manual testing

## instructions

Use this expert when writing automated integration tests that drive a bot's conversation programmatically — without a browser. Key use cases:
- CI pipeline smoke tests for Teams bots
- Multi-turn conversation scenario testing
- Non-chat activity simulation (email, install, member-added)
- Testing notification bots with specific `from` identities (personas)
- Streaming bot assertion (full assembled response)
- Cross-language test runners (Python, curl) via `ConversationServer`

Pair with `dev.debug-test-ts.md` for manual local debugging workflows and `toolkit.playground-ts.md` for the browser-based Agents Playground.

## research

Deep Research prompt:

"Write a micro expert on `@microsoft/m365agentsplayground-cli` for automated integration testing of Microsoft Teams bots (TypeScript). Cover: TestClient setup (botEndpoint, timeout, deliveryMode, start/stop), sendMessage() return value (BotResponse[]), newConversation() for test isolation, ConversationServer HTTP API (POST /run-conversation, GET /health, config schema, turn schema), turn types (chat, sendEmail, mentionInWord, install, userAdded, botAdded, channelCreated, teamRenamed), personas for notification activities, turn statuses (Completed, TimedOut, Errored, Skipped), streaming bot support (streamType:final, streamingSettleDelayMs), parallel testing (multiple TestClient instances, multiple ConversationServer ports), chat types (personal, group, channel), installation via npm or build from monorepo. Include canonical patterns for: basic Mocha test, Python POST /run-conversation, mixed turn types in one conversation, bot install event test."
