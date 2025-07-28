# Microsoft 365 Agents Toolkit - Template Dependency Analysis

**Analysis Date:** July 28, 2025  
**Repository:** microsoft-365-agents-toolkit  
**Directory Analyzed:** `templates/vsc/js/`

## Overview

This document provides a comprehensive analysis of all JavaScript templates in the Microsoft 365 Agents Toolkit, focusing on their dependencies and architectural patterns. The analysis reveals three distinct approaches to building Microsoft 365 applications within the toolkit.

## Template Categories

### 1. Microsoft Agents SDK Templates
Templates using the new `@microsoft/agents-*` packages represent the latest approach to building Microsoft 365 applications.

### 2. Teams AI Library Templates  
Templates using `@microsoft/teams-ai` for AI-powered bots and copilots.

### 3. Other Framework Templates
Templates using Azure Functions, React, or other specialized frameworks.

---



## Detailed Analysis by Framework

### Microsoft Agents SDK Templates (11 templates)

The newest approach using `@microsoft/agents-*` packages:

#### Core Microsoft Agents SDK Templates
| Template | Hosting Packages | Specialized Dependencies |
|----------|------------------|-------------------------|
| **basic-custom-engine-agent** | `@microsoft/agents-hosting` ^0.2.14 | `@azure/identity` ^4.8.0, `@azure/openai` ^2.0.0, `openai` ^4.94.0 |
| **default-bot** | `@microsoft/agents-hosting` ^0.2.14 | None (minimal bot) |
| **weather-agent** | `@microsoft/agents-hosting` ^0.2.14 | `@langchain/langgraph` ^0.2.66, `@azure/openai` ^2.0.0, `zod` 3.25.67 |

#### Message Extension Templates (Microsoft Agents SDK)
| Template | Dependencies |
|----------|-------------|
| **link-unfurling** | `@microsoft/agents-hosting` ^0.2.14, `@microsoft/agents-hosting-teams` ^0.2.14 |
| **m365-message-extension** | `@microsoft/agents-hosting` ^0.2.14, `@microsoft/agents-hosting-teams` ^0.2.14, `adaptivecards` ^3.0.1 |
| **message-extension** | `@microsoft/agents-hosting` ^0.2.14, `@microsoft/agents-hosting-teams` ^0.2.14, `isomorphic-fetch` ^3.0.0 |
| **message-extension-action** | `@microsoft/agents-hosting` ^0.2.14, `@microsoft/agents-hosting-teams` ^0.2.14, `adaptivecards` ^3.0.1 |

#### Notification Templates (Microsoft Agents SDK)
All notification templates use version ^0.4.3 of the hosting packages:

| Template | Trigger Type | Dependencies |
|----------|-------------|-------------|
| **notification-express** | HTTP (Express) | `@microsoft/agents-hosting` ^0.4.3, `@microsoft/agents-hosting-teams` ^0.4.3, `@microsoft/agents-activity` ^0.4.3 |
| **notification-http-timer-trigger** | HTTP + Timer (Azure Functions) | `@microsoft/agents-hosting` ^0.4.3, `@microsoft/agents-hosting-teams` ^0.4.3, `@microsoft/agents-activity` ^0.4.3 |
| **notification-http-trigger** | HTTP (Azure Functions) | `@microsoft/agents-hosting` ^0.4.3, `@microsoft/agents-hosting-teams` ^0.4.3, `@microsoft/agents-activity` ^0.4.3 |
| **notification-timer-trigger** | Timer (Azure Functions) | `@microsoft/agents-hosting` ^0.4.3, `@microsoft/agents-hosting-teams` ^0.4.3, `@microsoft/agents-activity` ^0.4.3 |

**Key Observations:**
- Notification templates use newer versions (^0.4.3) compared to other Agents SDK templates (^0.2.14)
- All include `@microsoft/agents-activity` for activity handling
- Express-based vs Azure Functions-based hosting options

---

### Teams AI Library Templates (9 templates)

Templates using `@microsoft/teams-ai` for AI-powered functionality:

#### Version Distribution
| Version | Templates | Notes |
|---------|-----------|-------|
| **^1.7.0** | command-and-response, workflow | Latest version |
| **^1.6.1** | custom-copilot-assistant-assistants-api | Pinned version (~1.6.1) |
| **^1.5.3** | custom-copilot-assistant-new, custom-copilot-basic, custom-copilot-rag-azure-ai-search, custom-copilot-rag-customize, custom-copilot-rag-microsoft365 | Most common version |
| **^1.1.0** | custom-copilot-rag-custom-api | Oldest version |

#### AI Assistant Templates
| Template | Teams AI Version | Specialization |
|----------|------------------|----------------|
| **custom-copilot-assistant-assistants-api** | ~1.6.1 | OpenAI Assistants API integration |
| **custom-copilot-assistant-new** | ^1.5.3 | General AI assistant |
| **custom-copilot-basic** | ^1.5.3 | Basic AI chat functionality |

#### RAG (Retrieval-Augmented Generation) Templates
| Template | Teams AI Version | Data Source | Additional Dependencies |
|----------|------------------|-------------|------------------------|
| **custom-copilot-rag-azure-ai-search** | ^1.5.3 | Azure AI Search | `@azure/search-documents` ^12.0.0 |
| **custom-copilot-rag-custom-api** | ^1.1.0 | Custom API | `openapi-client-axios` ^7.4.0, `js-yaml` ^4.1.0 |
| **custom-copilot-rag-customize** | ^1.5.3 | Custom data source | `@azure/search-documents` ^12.0.0 |
| **custom-copilot-rag-microsoft365** | ^1.5.3 | Microsoft 365 data | `@microsoft/microsoft-graph-client` ^3.0.1, `@azure/search-documents` ^12.0.0 |

#### Bot Framework Templates
| Template | Teams AI Version | Purpose |
|----------|------------------|---------|
| **command-and-response** | ^1.7.0 | Command handling bot |
| **workflow** | ^1.7.0 | Workflow automation bot |

**Common Dependencies:**
- All Teams AI templates use `botbuilder` ^4.23.1
- Most use `express` ^5.0.1
- Adaptive Cards templates include `adaptivecards-templating` ^2.3.1

---

## Complete Template Inventory

| Template Name | Primary Framework | Key Dependencies |
|---------------|-------------------|------------------|
| **basic-custom-engine-agent** | Microsoft Agents SDK | `@microsoft/agents-hosting` ^0.2.14, `@azure/identity` ^4.8.0, `@azure/openai` ^2.0.0, `openai` ^4.94.0 |
| **command-and-response** | Teams AI Library | `@microsoft/teams-ai` ^1.7.0, `botbuilder` ^4.23.1, `adaptive-expressions` ^4.23.1 |
| **custom-copilot-assistant-assistants-api** | Teams AI Library | `@microsoft/teams-ai` ~1.6.1, `botbuilder` ^4.23.1 |
| **custom-copilot-assistant-new** | Teams AI Library | `@microsoft/teams-ai` ^1.5.3, `botbuilder` ^4.23.1 |
| **custom-copilot-basic** | Teams AI Library | `@microsoft/teams-ai` ^1.5.3, `botbuilder` ^4.23.1 |
| **custom-copilot-rag-azure-ai-search** | Teams AI Library | `@microsoft/teams-ai` ^1.5.3, `@azure/search-documents` ^12.0.0, `botbuilder` ^4.23.1 |
| **custom-copilot-rag-custom-api** | Teams AI Library | `@microsoft/teams-ai` ^1.1.0, `botbuilder` ^4.23.1, `openapi-client-axios` ^7.4.0 |
| **custom-copilot-rag-customize** | Teams AI Library | `@microsoft/teams-ai` ^1.5.3, `@azure/search-documents` ^12.0.0, `botbuilder` ^4.23.1 |
| **custom-copilot-rag-microsoft365** | Teams AI Library | `@microsoft/teams-ai` ^1.5.3, `@microsoft/microsoft-graph-client` ^3.0.1, `@azure/search-documents` ^12.0.0 |
| **dashboard-tab** | React/Vite | `@fluentui/react-components` ^9.55.1, `@microsoft/teams-js` ^2.31.1, `react` ^18.2.0 |
| **declarative-agent-with-action-from-scratch** | Azure Functions | `@azure/functions` ^4.3.0 |
| **declarative-agent-with-action-from-scratch-bearer** | Azure Functions | `@azure/functions` ^4.3.0 |
| **declarative-agent-with-action-from-scratch-oauth** | Azure Functions | `@azure/functions` ^4.3.0, `jsonwebtoken` ^9.0.2, `jwks-rsa` ^3.1.0 |
| **default-bot** | Microsoft Agents SDK | `@microsoft/agents-hosting` ^0.2.14 |
| **default-bot-message-extension** | Bot Builder | `botbuilder` ^4.23.1, `adaptivecards` ^3.0.1 |
| **link-unfurling** | Microsoft Agents SDK | `@microsoft/agents-hosting` ^0.2.14, `@microsoft/agents-hosting-teams` ^0.2.14 |
| **m365-message-extension** | Microsoft Agents SDK | `@microsoft/agents-hosting` ^0.2.14, `@microsoft/agents-hosting-teams` ^0.2.14 |
| **message-extension** | Microsoft Agents SDK | `@microsoft/agents-hosting` ^0.2.14, `@microsoft/agents-hosting-teams` ^0.2.14 |
| **message-extension-action** | Microsoft Agents SDK | `@microsoft/agents-hosting` ^0.2.14, `@microsoft/agents-hosting-teams` ^0.2.14 |
| **message-extension-with-api-from-scratch** | Azure Functions | `@azure/functions` ^4.3.0 |
| **message-extension-with-api-from-scratch-api-key** | Azure Functions | `@azure/functions` ^4.3.0 |
| **message-extension-with-api-from-scratch-sso** | Azure Functions | `@azure/functions` ^4.3.0, `jsonwebtoken` ^9.0.2, `jwks-rsa` ^3.1.0 |
| **non-sso-tab** | Express Static | `express` ^4.21.1, `send` ^0.18.0 |
| **non-sso-tab-default-bot** | Multi-component | `concurrently` ^7.6.0 (separate bot and tab packages) |
| **notification-express** | Microsoft Agents SDK | `@microsoft/agents-hosting` ^0.4.3, `@microsoft/agents-activity` ^0.4.3 |
| **notification-http-timer-trigger** | Microsoft Agents SDK | `@microsoft/agents-hosting` ^0.4.3, `@microsoft/agents-activity` ^0.4.3 |
| **notification-http-trigger** | Microsoft Agents SDK | `@microsoft/agents-hosting` ^0.4.3, `@microsoft/agents-activity` ^0.4.3 |
| **notification-timer-trigger** | Microsoft Agents SDK | `@microsoft/agents-hosting` ^0.4.3, `@microsoft/agents-activity` ^0.4.3 |
| **sso-tab-naa** | React/Vite | `@azure/msal-browser` ^4.12.0, `@microsoft/teams-js` ^2.31.1, `react` ^18.2.0 |
| **weather-agent** | Microsoft Agents SDK | `@microsoft/agents-hosting` ^0.2.14, `@langchain/langgraph` ^0.2.66, `@azure/openai` ^2.0.0 |
| **workflow** | Teams AI Library | `@microsoft/teams-ai` ^1.7.0, `botbuilder` ^4.23.1, `adaptive-expressions` ^4.23.1 |

---

### Other Framework Templates (12 templates)

#### Legacy/Bot Builder Templates (3 templates)
| Template | Framework | Purpose |
|----------|-----------|---------|
| **default-bot-message-extension** | Bot Builder v4 | Message extension with traditional Bot Framework |
| **non-sso-tab-default-bot** | Multi-component | Combination of tab and bot in separate packages |

**Note:** Office add-in templates (`office-json-addin`, `office-xml-addin-*`) don't have `package.json.tpl` files and likely use different templating mechanisms.

---

## Architecture Patterns

### 1. Microsoft Agents SDK Pattern
- **Hosting:** `@microsoft/agents-hosting` + `@microsoft/agents-hosting-teams`
- **Activity Management:** `@microsoft/agents-activity` (for notifications)
- **Benefits:** Streamlined development, built-in Teams integration
- **Use Cases:** Modern Microsoft 365 applications, message extensions, notifications

### 2. Teams AI Library Pattern
- **Core:** `@microsoft/teams-ai` + `botbuilder`
- **AI Integration:** Built-in AI orchestration and prompt management
- **Benefits:** Advanced AI capabilities, RAG support, conversation management
- **Use Cases:** AI-powered bots, copilots, intelligent assistants


## Version Analysis

### Microsoft Agents SDK Versions
- **Core templates:** ^0.2.14
- **Notification templates:** ^0.4.3 (newer)
- **Trend:** Active development with version progression

### Teams AI Library Versions
- **Latest:** ^1.7.0 (command-and-response, workflow)
- **Stable:** ^1.5.3 (most copilot templates)
- **Legacy:** ^1.1.0 (custom-copilot-rag-custom-api)
- **Trend:** Most templates are on recent stable versions

### Common Dependencies
- **Express:** Ranges from ^4.21.1 to ^5.0.1
- **Bot Builder:** Consistently ^4.23.1
- **Adaptive Cards:** ^3.0.1 to ^3.0.5
- **Node.js:** All templates support "18 || 20 || 22"

---

## Development Patterns

### Script Patterns
Common across most templates:
```json
{
  "dev:teamsfx": "env-cmd --silent -f .localConfigs npm run dev",
  "dev:teamsfx:testtool": "env-cmd --silent -f .localConfigs.playground npm run dev",
  "dev:teamsfx:launch-testtool": "env-cmd --silent -f env/.env.playground teamsapptester start"
}
```

### Development Dependencies
Standard across templates:
- `env-cmd` ^10.1.0 (environment configuration)
- `nodemon` ^3.1.7 (development server)
- `concurrently` (for multi-component templates)

---

## Recommendations

### For New Projects
1. **Choose Microsoft Agents SDK** for new Microsoft 365 applications
2. **Use Teams AI Library** for AI-powered conversational experiences
3. **Consider Azure Functions** for lightweight API backends
4. **Use React patterns** for rich interactive tabs

### Migration Considerations
- **From Bot Builder:** Consider Teams AI Library for AI features
- **To Agents SDK:** Evaluate benefits of streamlined hosting
- **Version Updates:** Keep AI Library templates on latest versions

### Architecture Decision Factors
- **Complexity:** Agents SDK for simple, Functions for minimal
- **AI Requirements:** Teams AI Library for advanced AI features
- **UI Needs:** React patterns for rich interfaces
- **Hosting Preferences:** Express vs Azure Functions vs Agents hosting

---

## Conclusion

The Microsoft 365 Agents Toolkit demonstrates a strategic evolution in development approaches:

1. **Microsoft Agents SDK** represents the future direction for Microsoft 365 application development
2. **Teams AI Library** provides sophisticated AI capabilities for conversational experiences  
3. **Azure Functions** offers lightweight, serverless options for API development
4. **React/Tab templates** enable rich, modern user interfaces

The diversity of templates allows developers to choose the most appropriate technology stack based on their specific requirements, while the consistent patterns across templates ensure a cohesive development experience.
