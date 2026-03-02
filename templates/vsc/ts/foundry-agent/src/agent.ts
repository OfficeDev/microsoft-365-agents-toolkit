import { AIProjectClient } from "@azure/ai-projects";
import {
  AzureCliCredential,
  AzureDeveloperCliCredential,
  ChainedTokenCredential,
  ManagedIdentityCredential,
  TokenCredential,
} from "@azure/identity";
import { ActivityTypes } from "@microsoft/agents-activity";
import { AgentApplication, MemoryStorage, TurnContext } from "@microsoft/agents-hosting";
import config from "./config";

// Initialize credential based on environment
const environment = process.env.RUNNING_ON_AZURE === "1" ? "production" : "development";
const managedIdentityClientId = process.env.MI_CLIENT_ID;
let credential: TokenCredential;

console.log(`[INIT] Environment: ${environment}`);
console.log(`[INIT] NODE_ENV: ${process.env.NODE_ENV}`);
console.log(`[INIT] FOUNDRY_PROJECT_ENDPOINT: ${process.env.FOUNDRY_PROJECT_ENDPOINT}`);

if (environment === "production") {
  console.log(`[INIT] Using ManagedIdentityCredential with clientId=${managedIdentityClientId}`);
  credential = new ManagedIdentityCredential({
    clientId: managedIdentityClientId,
  });
} else {
  // Local development: Use ChainedTokenCredential for explicit, predictable behavior
  // This avoids the "fail fast" mode issues with DefaultAzureCredential
  console.log(
    "[INIT] Development environment: Using ChainedTokenCredential (AzureCli -> AzureDeveloperCli)"
  );

  credential = new ChainedTokenCredential(
    new AzureCliCredential(),
    new AzureDeveloperCliCredential()
  );
}

// Initialize Microsoft Foundry (Azure AI Projects) client
const projectClient = new AIProjectClient(process.env.FOUNDRY_PROJECT_ENDPOINT || "", credential);

console.log("Azure AI Agent Service client initialized successfully");

// Helper function for retry logic with exponential backoff
async function retryWithBackoff<T>(
  fn: () => Promise<T>,
  maxRetries = 3,
  initialDelayMs = 1000
): Promise<T> {
  let lastError: any;

  for (let attempt = 0; attempt < maxRetries; attempt++) {
    try {
      return await fn();
    } catch (error) {
      lastError = error;
      if (attempt < maxRetries - 1) {
        const delayMs = initialDelayMs * Math.pow(2, attempt);
        console.log(`[RETRY] Attempt ${attempt + 1} failed, retrying in ${delayMs}ms...`);
        await new Promise((resolve) => setTimeout(resolve, delayMs));
      }
    }
  }

  throw lastError;
}

// Helper function to create an OAuth consent card
function createOAuthConsentCard(consentLink: string, serviceName: string): any {
  return {
    $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
    type: "AdaptiveCard",
    version: "1.4",
    body: [
      {
        type: "TextBlock",
        text: "Authorization Required",
        weight: "bolder",
        size: "large",
      },
      {
        type: "TextBlock",
        text: `The agent needs authorization to access ${serviceName}. Click the button below to grant permission.`,
        wrap: true,
        spacing: "medium",
      },
      {
        type: "TextBlock",
        text: "1. Click 'Authorize' to open the authorization page",
        wrap: true,
        spacing: "small",
        size: "small",
      },
      {
        type: "TextBlock",
        text: "2. Complete the authorization process",
        wrap: true,
        spacing: "small",
        size: "small",
      },
      {
        type: "TextBlock",
        text: "3. Return to Teams and send 'retry' to continue",
        wrap: true,
        spacing: "small",
        size: "small",
      },
      {
        type: "TextBlock",
        text: "You will be redirected to authorize the application.",
        wrap: true,
        spacing: "medium",
        size: "small",
        color: "warning",
      },
    ],
    actions: [
      {
        type: "Action.OpenUrl",
        title: `Authorize ${serviceName}`,
        url: consentLink,
        style: "positive",
      },
    ],
  };
}

function createApprovalCard(
  requestName: string,
  requestArgs: any,
  requestId: string,
  conversationId: string
): any {
  return {
    $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
    type: "AdaptiveCard",
    version: "1.4",
    body: [
      {
        type: "TextBlock",
        text: "Action Approval Required",
        weight: "bolder",
        size: "large",
      },
      {
        type: "TextBlock",
        text: `The agent needs your approval to perform: ${requestName}`,
        wrap: true,
        spacing: "medium",
      },
      {
        type: "TextBlock",
        text: "Details:",
        weight: "bolder",
        spacing: "small",
      },
      {
        type: "TextBlock",
        text: JSON.stringify(requestArgs, null, 2),
        wrap: true,
        fontType: "monospace",
        size: "small",
      },
    ],
    actions: [
      {
        type: "Action.Submit",
        title: "✓ Approve",
        style: "positive",
        disabled: false,
        data: {
          action: "approve",
          requestId: requestId,
          conversationId: conversationId,
        },
      },
      {
        type: "Action.Submit",
        title: "✗ Deny",
        style: "destructive",
        disabled: false,
        data: {
          action: "deny",
          requestId: requestId,
          conversationId: conversationId,
        },
      },
    ],
  };
}

// Define storage and application
const storage = new MemoryStorage();

// Store pending conversations that are waiting for OAuth
const pendingOAuthConversations = new Map<
  string,
  {
    conversationId: string;
    originalMessage: string;
    timestamp: number;
  }
>();

export const agentApp = new AgentApplication({
  storage,
});

// Helper function to handle approval card submissions
async function handleApprovalResponse(context: TurnContext, value: any): Promise<boolean> {
  if (!value || !value.action || !value.requestId) {
    return false;
  }

  console.log(
    `[APPROVAL HANDLER] User ${value.action === "approve" ? "approved" : "denied"} request: ${
      value.requestId
    }`
  );

  const actionLabel = value.action === "approve" ? "✓ Approved" : "✗ Denied";
  await context.sendActivity(`${actionLabel} - Processing your request...`);

  const openAIClient = await projectClient.getOpenAIClient();
  const approved = value.action === "approve";
  const conversationId = value.conversationId as string;
  const requestId = value.requestId as string;

  console.log(`[APPROVAL HANDLER] Conversation ID: ${conversationId}`);

  try {
    console.log(`[APPROVAL HANDLER] Processing approval/denial for request ${requestId}`);

    const approvalResponse = await retryWithBackoff(
      () =>
        openAIClient.responses.create(
          {
            conversation: conversationId,
            input: [
              {
                type: "mcp_approval_response",
                approve: approved,
                approval_request_id: requestId,
              },
            ],
          },
          {
            body: {
              agent: {
                name: config.foundryAgentName,
                type: "agent_reference",
              },
            },
          }
        ),
      3,
      1000
    );

    const resultMessage = approved
      ? `${(approvalResponse as any).output_text || "Processing your request..."}`
      : `✗ Action denied.`;

    console.log(`[APPROVAL HANDLER] Approval processed successfully`);
    console.log(
      `[APPROVAL HANDLER] Response:`,
      (approvalResponse as any).output_text?.substring(0, 100)
    );
    await context.sendActivity(resultMessage);
  } catch (approvalError) {
    console.error(
      "[APPROVAL HANDLER] Error message:",
      (approvalError as any).error?.message || (approvalError as any).message
    );
    console.error("[APPROVAL HANDLER] Full error:", {
      message: (approvalError as any).message,
      code: (approvalError as any).code,
      status: (approvalError as any).status,
    });

    const fallbackMessage = approved ? `✓ Action approved!` : `✗ Action denied.`;

    console.log(`[APPROVAL HANDLER] Falling back to simple acknowledgment`);
    await context.sendActivity(fallbackMessage);
  }

  return true;
}

// Helper function to handle OAuth consent requests in response
async function handleOAuthConsentRequest(
  context: TurnContext,
  response: any,
  conversationId: string,
  originalMessage: string
): Promise<boolean> {
  const oauthRequest = response.output?.find((item: any) => item.type === "oauth_consent_request");

  if (!oauthRequest) {
    return false;
  }

  console.log(`[OAUTH HANDLER] Found OAuth consent request in response output`);

  const consentLink = (oauthRequest as any).consent_link || "";
  const serviceName = (oauthRequest as any).service_name || "Service";

  if (!consentLink) {
    console.error(`[OAUTH HANDLER] OAuth consent request missing consent_link`);
    await context.sendActivity(
      "⚠️ Authorization is required, but the authorization link is missing. " +
        "Please authorize the connections manually in Azure AI Foundry Studio."
    );
    return true;
  }

  const oauthCard = createOAuthConsentCard(consentLink, serviceName);

  // Store this conversation in pending OAuth map for retry
  const conversationKey = `${context.activity.channelId}:${context.activity.from.id}`;
  pendingOAuthConversations.set(conversationKey, {
    conversationId,
    originalMessage,
    timestamp: Date.now(),
  });
  console.log(`[OAUTH HANDLER] Stored pending OAuth conversation for key: ${conversationKey}`);

  await context.sendActivity({
    type: ActivityTypes.Message,
    attachments: [
      {
        contentType: "application/vnd.microsoft.card.adaptive",
        content: oauthCard,
      },
    ],
  } as any);

  console.log(`[OAUTH HANDLER] Sent OAuth consent card, waiting for user authorization`);
  return true;
}

// Helper function to handle approval requests in response
async function handleApprovalRequest(
  context: TurnContext,
  response: any,
  conversationId: string
): Promise<boolean> {
  const approvalRequest = response.output?.find(
    (item: any) => item.type === "mcp_approval_request"
  );

  if (!approvalRequest) {
    return false;
  }

  console.log(`[APPROVAL REQUEST] Found approval request: ${(approvalRequest as any).id}`);

  const requestName = (approvalRequest as any).name || "Unknown Action";
  const requestArgs = (approvalRequest as any).arguments || {};
  const requestId = (approvalRequest as any).id || "";

  console.log(`[APPROVAL REQUEST] Details:`, JSON.stringify(approvalRequest, null, 2));

  const approvalCard = createApprovalCard(requestName, requestArgs, requestId, conversationId);

  await context.sendActivity({
    type: ActivityTypes.Message,
    attachments: [
      {
        contentType: "application/vnd.microsoft.card.adaptive",
        content: approvalCard,
      },
    ],
  } as any);

  console.log(`[APPROVAL REQUEST] Sent approval card, waiting for user response`);
  return true;
}

// Helper function to send regular text response
async function sendTextResponse(context: TurnContext, response: any): Promise<void> {
  const answer = response.output_text || "I'm sorry, I couldn't generate a response.";
  console.log(`[RESPONSE] Sending text response: ${answer.substring(0, 100)}...`);
  await context.sendActivity(answer);
}

agentApp.onConversationUpdate("membersAdded", async (context: TurnContext) => {
  await context.sendActivity(`Hi there! I'm an agent to chat with you.`);
});

// Listen for ANY message to be received. MUST BE AFTER ANY OTHER MESSAGE HANDLERS
agentApp.onActivity(ActivityTypes.Message, async (context: TurnContext) => {
  console.log("[MESSAGE HANDLER] Receive message");
  try {
    // Check if this is an adaptive card submission (approval response)
    const value = context.activity.value as any;
    const isApprovalResponse = await handleApprovalResponse(context, value);
    if (isApprovalResponse) {
      return; // Exit after handling approval
    }

    // Handle regular text message
    if (!context.activity.text) {
      console.log("[MESSAGE HANDLER] No text content, skipping");
      return;
    }

    const userMessage = context.activity.text.toLowerCase().trim();
    const conversationKey = `${context.activity.channelId}:${context.activity.from.id}`;

    // Check if user is retrying after OAuth
    if (userMessage === "retry" && pendingOAuthConversations.has(conversationKey)) {
      console.log("[MESSAGE HANDLER] User retrying after OAuth authorization");
      const pending = pendingOAuthConversations.get(conversationKey)!;
      pendingOAuthConversations.delete(conversationKey);

      try {
        const openAIClient = await projectClient.getOpenAIClient();

        // Retry the original request with the now-authorized credentials
        console.log(
          `[MESSAGE HANDLER] Retrying agent request for conversation: ${pending.conversationId}`
        );
        const retryResponse = await retryWithBackoff(
          () =>
            openAIClient.responses.create(
              {
                conversation: pending.conversationId,
              },
              {
                body: {
                  agent: {
                    name: config.foundryAgentName,
                    type: "agent_reference",
                  },
                },
              }
            ),
          3,
          1000
        );

        console.log(
          `[MESSAGE HANDLER] Retry response received with ${
            retryResponse.output?.length || 0
          } items`
        );

        // Send the successful response to the user
        const answer = retryResponse.output_text || "I'm sorry, I couldn't generate a response.";
        await context.sendActivity(answer);
        return;
      } catch (retryError) {
        console.error("[MESSAGE HANDLER] Retry failed:", retryError);
        await context.sendActivity("Failed to retry the request. Please try again.");
        return;
      }
    }

    console.log(`[MESSAGE HANDLER] Processing message: ${context.activity.text}`);
    console.log(`[MESSAGE HANDLER] Using Agent: ${config.foundryAgentName}`);

    // Print User Context and Scope
    console.log("========== USER CONTEXT START ==========");
    console.log(`User ID: ${context.activity.from.id}`);
    console.log(`User Name: ${context.activity.from.name}`);
    console.log(`AAD Object ID: ${context.activity.from.aadObjectId}`);
    console.log(`Channel ID: ${context.activity.channelId}`);
    try {
      console.log(
        `Tenant ID: ${
          context.activity.conversation.tenantId ||
          (context.activity.channelData as any)?.tenant?.id
        }`
      );
    } catch (e) {
      console.log("Tenant ID: unavailable");
    }
    console.log("========== USER CONTEXT END ==========");

    // Test credential by getting a token first
    console.log(`[MESSAGE HANDLER] Testing credential...`);
    const tokenScope = "https://cognitiveservices.azure.com/.default";
    console.log(`[MESSAGE HANDLER] Requesting Token Scope: ${tokenScope}`);

    try {
      const token = await credential.getToken(tokenScope);
      console.log(
        `[MESSAGE HANDLER] Token obtained successfully, expires: ${token?.expiresOnTimestamp}`
      );
    } catch (tokenError) {
      console.error(`[MESSAGE HANDLER] Failed to get token:`, tokenError);
      throw tokenError;
    }

    // Get the extended OpenAI client from project (with agent APIs)
    console.log(`[MESSAGE HANDLER] Getting OpenAI client from project...`);
    const openAIClient = await projectClient.getOpenAIClient();
    console.log(`[MESSAGE HANDLER] OpenAI client obtained successfully`);

    // Create a conversation with the user's message
    console.log(`[MESSAGE HANDLER] Creating conversation...`);
    const conversation = await openAIClient.conversations.create({
      items: [{ type: "message", role: "user", content: context.activity.text }],
    });

    console.log(`[MESSAGE HANDLER] Created conversation: ${conversation.id}`);

    // Generate response using the foundry agent reference
    console.log(`[MESSAGE HANDLER] Generating response with agent...`);
    let response: any;
    try {
      response = await retryWithBackoff(
        () =>
          openAIClient.responses.create(
            {
              conversation: conversation.id,
            },
            {
              body: {
                agent: {
                  name: config.foundryAgentName,
                  type: "agent_reference",
                },
              },
            }
          ),
        3,
        1000
      );
      console.log(`[MESSAGE HANDLER] Response created successfully`);
    } catch (responseError: any) {
      console.error(`[MESSAGE HANDLER] Error creating response:`, {
        message: responseError?.message,
        code: responseError?.code,
        type: responseError?.type,
        status: responseError?.status,
      });

      // Check if this is an OAuth-related error
      if (
        responseError?.message?.includes("Failed to fetch access token") ||
        responseError?.code === "tool_user_error"
      ) {
        console.log(`[MESSAGE HANDLER] Detected OAuth/connection authorization error`);
        console.log(
          `[MESSAGE HANDLER] This typically means a connection in the agent needs user authorization`
        );

        // Send a helpful message to the user explaining the issue
        await context.sendActivity(
          "⚠️ **Authorization Required**\n\n" +
            "The agent needs authorization to access data to help you with your tasks.\n\n" +
            "**How to fix this:**\n" +
            "1. Go to [Azure AI Foundry Studio](https://ai.azure.com)\n" +
            "2. Navigate to your project: `agent-dev-project`\n" +
            "3. Go to **Connections** section\n" +
            "4. Find the **MicrosoftSharePointand** connection\n" +
            "5. Click **Authorize** and sign in with your Microsoft 365 account\n\n" +
            "After authorizing, try your request again!"
        );
        return;
      }

      // If it's a different error, re-throw it
      throw responseError;
    }

    console.log(`[MESSAGE HANDLER] Got response with ${response.output?.length || 0} output items`);

    // Check if response contains special requests (OAuth, approval, etc.)
    console.log(`[MESSAGE HANDLER] Checking response output for special requests...`);
    if (response.output && Array.isArray(response.output)) {
      console.log(
        `[MESSAGE HANDLER] Response output types:`,
        response.output.map((item: any) => item.type)
      );

      // Check for OAuth consent requests first
      const hasOAuthRequest = await handleOAuthConsentRequest(
        context,
        response,
        conversation.id,
        context.activity.text || ""
      );
      if (hasOAuthRequest) {
        return; // Wait for user to authorize
      }

      // Check for approval requests
      const hasApprovalRequest = await handleApprovalRequest(context, response, conversation.id);
      if (hasApprovalRequest) {
        return; // Wait for user approval
      }
    }

    // Send the text response back to the user
    await sendTextResponse(context, response);
  } catch (error) {
    console.error("[MESSAGE HANDLER] Error calling foundry agent:", {
      error: error,
      message: (error as any).message,
      stack: (error as any).stack,
    });
    await context.sendActivity("I'm sorry, I encountered an error while processing your request.");
  }
});
