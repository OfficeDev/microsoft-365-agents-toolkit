# Microsoft 365 Agents Toolkit - Deployment Guide

## Overview

This guide covers deployment options for Microsoft 365 agents and applications built with the Microsoft 365 Agents Toolkit. This toolkit is specifically designed for building:
- Microsoft 365 Copilot agents
- Microsoft Teams applications
- Office Add-ins
- Microsoft Agents SDK-based solutions

> **Note**: This toolkit is NOT for deploying ML models, FastAPI applications, or general-purpose web applications. For ML deployment, please refer to appropriate ML deployment platforms.

## Deployment Methods

### 1. Azure Deployment (Recommended)

The toolkit provides built-in Azure deployment capabilities:

```bash
# Deploy to Azure using CLI
m365agents deploy --env <environment>
```

This command:
- Deploys your application to Azure App Service
- Configures necessary Azure resources
- Sets up bot registrations and authentication
- Updates environment configurations

### 2. CI/CD Deployment

The toolkit supports automated deployment through:
- GitHub Actions workflows
- Azure DevOps pipelines

Configuration files are automatically generated in your project templates.

### 3. Local Development and Testing

For local testing before deployment:

```bash
# Install Microsoft 365 Agents Toolkit CLI
npm install -g @microsoft/m365agentstoolkit-cli

# Preview your app locally
m365agents preview --env local
```

## Deployment Configuration

Deployment settings are managed through:
- `m365agents.yml` - Main configuration file
- `m365agents.local.yml` - Local development configuration
- `m365agents.sandbox.yml` - Sandbox environment configuration

These files are located in your project templates under `/templates/vsc/` directory.

## Supported Platforms

The toolkit officially supports deployment to:
- **Azure App Service** - Primary hosting platform
- **Azure Functions** - For serverless scenarios
- **Teams App Catalog** - For Teams app distribution
- **Microsoft 365 Admin Center** - For Microsoft 365 apps

## Getting Help

For detailed deployment documentation, visit:
- [Microsoft 365 Agents Toolkit Documentation](https://aka.ms/teamsfx-docs)
- [Teams Toolkit Documentation](https://docs.microsoft.com/microsoftteams/platform/toolkit/teams-toolkit-fundamentals)
- [Bot Framework Troubleshooting](https://docs.microsoft.com/azure/bot-service/bot-service-troubleshoot-index)

## Not Supported

This toolkit does NOT provide:
- ML model deployment platforms
- Generic FastAPI backend deployment
- Railway/Render/Heroku deployment automation
- General-purpose web application hosting

For these use cases, please refer to appropriate tools and platforms designed for those purposes.
