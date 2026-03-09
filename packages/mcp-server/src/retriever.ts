// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { z } from "zod";

export const ResourceTypeEnum = z.enum(["documents", "samples", "issues", "code"]);
export type ResourceType = z.infer<typeof ResourceTypeEnum>;

/**
 * Static guidance documents for each resource type.
 * These instruct the AI to use the appropriate MCP servers for retrieving resources.
 */
const resourceDocuments: Record<ResourceType, string> = {
  documents: `To find Microsoft 365 agents and apps development documentation, use the Learn MCP server (https://learn.microsoft.com/api/mcp).

Recommended workflow:
1. Use the "microsoft_docs_search" tool to search for relevant Microsoft 365 and Copilot development documentation.
2. Use the "microsoft_docs_fetch" tool to fetch full content from specific documentation URLs when you need deeper details.

Key documentation areas:

## Microsoft 365 Copilot Extensibility
Build agents for Microsoft 365 Copilot including declarative agents, API plugins, MCP plugins, and custom engine agents.
- Overview: https://learn.microsoft.com/microsoft-365-copilot/extensibility/overview
- Agents overview: https://learn.microsoft.com/microsoft-365-copilot/extensibility/agents-overview
- Planning guide: https://learn.microsoft.com/microsoft-365-copilot/extensibility/planning-guide
- Declarative agents: https://learn.microsoft.com/microsoft-365-copilot/extensibility/overview-declarative-agent
- Build declarative agents: https://learn.microsoft.com/microsoft-365-copilot/extensibility/build-declarative-agents
- Declarative agent instructions: https://learn.microsoft.com/microsoft-365-copilot/extensibility/declarative-agent-instructions
- API plugins: https://learn.microsoft.com/microsoft-365-copilot/extensibility/overview-plugins
- Build API plugins: https://learn.microsoft.com/microsoft-365-copilot/extensibility/build-api-plugins-new-api
- MCP plugins: https://learn.microsoft.com/microsoft-365-copilot/extensibility/build-mcp-plugins
- Custom engine agents: https://learn.microsoft.com/microsoft-365-copilot/extensibility/overview-custom-engine-agent
- Copilot connectors: https://learn.microsoft.com/microsoft-365-copilot/extensibility/overview-copilot-connector
- Copilot APIs: https://learn.microsoft.com/microsoft-365-copilot/extensibility/copilot-apis-overview

### Agent Manifest References
- Declarative agent manifest v1.6: https://learn.microsoft.com/microsoft-365-copilot/extensibility/declarative-agent-manifest-1.6
- Plugin manifest v2.4: https://learn.microsoft.com/microsoft-365-copilot/extensibility/plugin-manifest-2.4
- OpenAPI spec guidance: https://learn.microsoft.com/microsoft-365-copilot/extensibility/openapi-document-guidance
- API plugin authentication: https://learn.microsoft.com/microsoft-365-copilot/extensibility/api-plugin-authentication

## Microsoft Teams Platform (Apps & Bots)
Build Teams apps with tabs, bots, message extensions, webhooks, and meeting extensions.
- Overview: https://learn.microsoft.com/microsoftteams/platform/overview
- Build agents with Teams SDK: https://learn.microsoft.com/microsoftteams/platform/agents-in-teams/overview
- Build your first agent: https://learn.microsoft.com/microsoftteams/platform/agents-in-teams/build-first-agent
- Build bots: https://learn.microsoft.com/microsoftteams/platform/bots/overview
- Build tabs: https://learn.microsoft.com/microsoftteams/platform/tabs/what-are-tabs
- Message extensions: https://learn.microsoft.com/microsoftteams/platform/messaging-extensions/what-are-messaging-extensions
- API-based message extensions: https://learn.microsoft.com/microsoftteams/platform/messaging-extensions/api-based-overview
- Apps for meetings: https://learn.microsoft.com/microsoftteams/platform/apps-in-teams-meetings/teams-apps-in-meetings
- Adaptive Cards: https://learn.microsoft.com/microsoftteams/platform/task-modules-and-cards/what-are-cards
- App manifest schema: https://learn.microsoft.com/microsoftteams/platform/resources/schema/manifest-schema
- Authentication (SSO): https://learn.microsoft.com/microsoftteams/platform/concepts/authentication/authentication
- Extend across Microsoft 365: https://learn.microsoft.com/microsoftteams/platform/m365-apps/overview
- MCP server connectors for agents: https://learn.microsoft.com/microsoftteams/platform/m365-apps/agent-connectors

## Microsoft 365 Agents Toolkit
The toolkit for building, debugging, and deploying Microsoft 365 agents and apps.
- Overview: https://learn.microsoft.com/microsoftteams/platform/toolkit/overview-agents-toolkit
- Agents Toolkit for VS Code: https://learn.microsoft.com/microsoftteams/platform/toolkit/agents-toolkit-fundamentals
- Install: https://learn.microsoft.com/microsoftteams/platform/toolkit/install-agents-toolkit
- Create new project: https://learn.microsoft.com/microsoftteams/platform/toolkit/create-new-project
- Debug overview: https://learn.microsoft.com/microsoftteams/platform/toolkit/debug-overview
- Debug with Agents Playground: https://learn.microsoft.com/microsoftteams/platform/toolkit/debug-your-agents-playground
- Provision cloud resources: https://learn.microsoft.com/microsoftteams/platform/toolkit/provision
- Deploy to cloud: https://learn.microsoft.com/microsoftteams/platform/toolkit/deploy
- Publish: https://learn.microsoft.com/microsoftteams/platform/toolkit/publish
- CI/CD pipelines: https://learn.microsoft.com/microsoftteams/platform/toolkit/use-CICD-template
- CLI: https://learn.microsoft.com/microsoftteams/platform/toolkit/microsoft-365-agents-toolkit-CLI

## Microsoft 365 Agents SDK
Build custom engine agents with multi-language SDK support.
- Overview: https://learn.microsoft.com/microsoft-365-copilot/extensibility/m365-agents-sdk
- Create and deploy: https://learn.microsoft.com/microsoft-365-copilot/extensibility/create-deploy-agents-sdk

## Other Key Resources
- Microsoft Graph API: https://learn.microsoft.com/graph/
- Bot Framework SDK: https://learn.microsoft.com/azure/bot-service/
- Adaptive Cards: https://learn.microsoft.com/adaptive-cards/
- Office Add-ins: https://learn.microsoft.com/office/dev/add-ins/overview/office-add-ins
- Teams Store validation: https://learn.microsoft.com/microsoftteams/platform/concepts/deploy-and-publish/appsource/prepare/teams-store-validation-guidelines
- Copilot validation guidelines: https://learn.microsoft.com/microsoftteams/platform/concepts/deploy-and-publish/appsource/prepare/review-copilot-validation-guidelines`,

  samples: `To find code samples for Microsoft 365 agents and apps development, use the Learn MCP server (https://learn.microsoft.com/api/mcp).

Recommended workflow:
1. Use the "microsoft_code_sample_search" tool to search for official Microsoft code samples. You can optionally filter by language (e.g., "typescript", "csharp", "python").
2. Use the "microsoft_docs_fetch" tool to fetch full sample details from a specific URL when needed.

Key sample resources:
- Copilot extensibility samples: https://learn.microsoft.com/microsoft-365-copilot/extensibility/samples
- Teams platform tutorials and code samples: https://learn.microsoft.com/microsoftteams/platform/get-started/tool-options-and-code-samples

Key sample repositories:

## Microsoft 365 Agents SDK
- microsoft/Agents: M365 Agents SDK samples (multi-language, includes Node.js samples under samples/**/nodejs/). https://github.com/microsoft/Agents
- microsoft/Agents-for-js: M365 Agents SDK for JavaScript samples. https://github.com/microsoft/Agents-for-js

## Microsoft Teams Samples & SDK
- OfficeDev/Microsoft-Teams-Samples: Official Microsoft Teams app samples covering bots, tabs, message extensions, meetings, and more (Node.js/JS/TS). https://github.com/OfficeDev/Microsoft-Teams-Samples
- microsoft/teams-ai: Teams SDK (formerly Teams AI library) JavaScript samples under js/samples/. https://github.com/microsoft/teams-ai

## Microsoft 365 Agents Toolkit
- OfficeDev/microsoft-365-agents-toolkit: Toolkit project templates (JS/TS) under templates/. https://github.com/OfficeDev/microsoft-365-agents-toolkit
- OfficeDev/microsoft-365-agents-toolkit-samples: Scenario-focused sample apps built with the Agents Toolkit. https://github.com/OfficeDev/microsoft-365-agents-toolkit-samples`,

  issues: `To find and troubleshoot issues related to Microsoft 365 agents and apps development, use the GitHub MCP server (https://api.githubcopilot.com/mcp/).

Recommended workflow:
1. Use the GitHub MCP server tools to search for issues in the relevant repositories listed below.
2. Search both open and closed issues for relevant error messages or keywords.
3. Check issue comments for workarounds and solutions.
4. When searching, exclude issues with labels like "cherry-pick-hotfix" or "invalid" as they are not relevant.

Repositories to search:

## Core Toolkit & Samples
- OfficeDev/microsoft-365-agents-toolkit: The core toolkit (VS Code extension, CLI, fx-core) for building, debugging, and deploying Microsoft 365 agents and apps. Exclude labels: cherry-pick-hotfix, invalid.
- OfficeDev/microsoft-365-agents-toolkit-samples: Scenario-focused sample apps built with the Agents Toolkit. Exclude labels: invalid.
- OfficeDev/Microsoft-Teams-Samples: Official Microsoft Teams app samples covering bots, tabs, message extensions, meetings, and more. Exclude labels: invalid.

## Teams SDK (formerly Teams AI library)
- microsoft/teams-sdk: Main repo for the Teams SDK with docs and cross-language issues. Exclude labels: invalid.
- microsoft/teams.ts: Teams SDK TypeScript implementation (language-specific bugs). Exclude labels: invalid.
- microsoft/teams.net: Teams SDK .NET implementation (language-specific bugs). Exclude labels: invalid.
- microsoft/teams.py: Teams SDK Python implementation (language-specific bugs). Exclude labels: invalid.
- microsoft/teams-agent-accelerator-templates: Open-source agent accelerator templates for the Teams SDK. Exclude labels: invalid.

## Documentation
- MicrosoftDocs/msteams-docs: Microsoft Teams platform documentation issues and feedback. Exclude labels: invalid.
- MicrosoftDocs/m365copilot-docs: Microsoft 365 Copilot extensibility documentation issues. Exclude labels: invalid.
- OfficeDev/office-js-docs-pr: Office Add-ins documentation issues. Exclude labels: invalid.

Also check these resources for known issues:
- Copilot extensibility known issues: https://learn.microsoft.com/microsoft-365-copilot/extensibility/known-issues
- Teams platform troubleshooting: https://learn.microsoft.com/microsoftteams/platform/resources/troubleshoot
- Teams Toolkit FAQ: https://learn.microsoft.com/microsoftteams/platform/toolkit/faq`,

  code: `To find code snippets and SDK usage examples for Microsoft 365 agents and apps development, use the Learn MCP server (https://learn.microsoft.com/api/mcp).

Recommended workflow:
1. Use the "microsoft_code_sample_search" tool to find implementation examples for SDKs such as @microsoft/teams-ai, @microsoft/teams-js, and botbuilder. You can optionally filter by language.
2. Use the "microsoft_docs_search" tool to find documentation pages with embedded code examples.
3. Use the "microsoft_docs_fetch" tool to retrieve full code examples from specific documentation URLs.

Key SDKs and references:

## Teams SDK (Teams AI Library)
AI-powered Teams bot and agent development.
- Getting started: https://learn.microsoft.com/microsoftteams/platform/teams-ai-library/welcome
- Build your first agent: https://learn.microsoft.com/microsoftteams/platform/agents-in-teams/build-first-agent
- TypeScript SDK reference: https://learn.microsoft.com/javascript/api/teams-sdk-typescript
- .NET SDK reference: https://learn.microsoft.com/dotnet/api/?view=msteams-sdk-dotnet-latest

## Microsoft 365 Agents SDK
Multi-language SDK for building custom engine agents.
- .NET SDK: https://learn.microsoft.com/dotnet/api/?view=m365-agents-sdk
- JavaScript SDK: https://learn.microsoft.com/javascript/api/overview/agents-overview?view=agents-sdk-js-latest
- Python SDK: https://learn.microsoft.com/python/api/agent-sdk-python/agents-overview?view=agent-sdk-python-latest

## Other SDKs
- @microsoft/teams-js (Teams JavaScript client library): https://learn.microsoft.com/microsoftteams/platform/tabs/how-to/using-teams-client-library
- TeamsFx SDK: https://learn.microsoft.com/microsoftteams/platform/toolkit/TeamsFx-SDK
- botbuilder (Bot Framework SDK for Node.js): https://github.com/Microsoft/botbuilder-js
- Microsoft.Bot.Builder (Bot Framework SDK for .NET): https://github.com/Microsoft/botbuilder-dotnet
- Microsoft Graph SDKs: https://learn.microsoft.com/graph/sdks/sdks-overview`,
};

/**
 * Retrieves static guidance documents for the given resource type.
 * The documents instruct the AI to use the Learn MCP server or GitHub MCP server
 * to fetch the actual resources.
 * @param resourceType The type of resource to retrieve guidance for
 * @param question The question from the user (included in the response for context)
 * @returns Static guidance document as a string
 */
export function retrieveResource(resourceType: ResourceType, question: string): string {
  const guidance = resourceDocuments[resourceType];
  return `User question: "${question}"\n\n${guidance}`;
}
