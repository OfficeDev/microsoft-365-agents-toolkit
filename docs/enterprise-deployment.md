# Enterprise Deployment Guide

## Overview

This guide provides comprehensive instructions for deploying Microsoft 365 Agents in enterprise environments with CI/CD automation, integration with enterprise tools, and production-ready configurations.

## Table of Contents

- [Prerequisites](#prerequisites)
- [GitHub Enterprise Integration](#github-enterprise-integration)
- [CI/CD Pipeline Setup](#cicd-pipeline-setup)
- [Atlassian Integration](#atlassian-integration)
- [Slack Integration](#slack-integration)
- [Security Best Practices](#security-best-practices)
- [Monitoring and Observability](#monitoring-and-observability)

## Prerequisites

### Required Tools
- Microsoft 365 Agents Toolkit CLI: `npm install -g @microsoft/m365agentstoolkit-cli`
- Node.js >= 22
- Azure CLI (for Azure deployments)
- Git (for version control)

### Required Accounts
- Microsoft 365 tenant with administrator access
- Azure subscription (for cloud hosting)
- GitHub account (for CI/CD)

## GitHub Enterprise Integration

### Setting Up GitHub Actions

1. **Create Repository Secrets**
   
   Navigate to your repository settings and add the following secrets:
   
   ```
   BOT_ID                    # Microsoft Bot ID
   BOT_PASSWORD              # Microsoft Bot Password
   AZURE_SUBSCRIPTION_ID     # Azure Subscription ID
   AZURE_RESOURCE_GROUP      # Azure Resource Group Name
   TEAMS_APP_ID              # Teams Application ID
   ```

2. **Configure Branch Protection**
   
   Set up branch protection rules for `main` and `production` branches:
   - Require pull request reviews
   - Require status checks to pass
   - Require linear history
   - Include administrators

3. **Create Deployment Environments**
   
   Set up GitHub Environments for different deployment stages:
   - `development` - For active development
   - `staging` - For pre-production testing
   - `production` - For live deployments

### Example Workflow Structure

See `.github/workflows/enterprise-deploy.yml.example` for a complete CI/CD workflow template.

## CI/CD Pipeline Setup

### Multi-Environment Deployment Strategy

```yaml
Environments:
  development:  # Feature development and testing
    trigger: push to feature branches
    approval: automatic
    
  staging:      # Pre-production validation
    trigger: merge to main
    approval: automatic
    
  production:   # Live deployment
    trigger: manual or tag release
    approval: required (designated approvers)
```

### Deployment Steps

1. **Build Phase**
   - Install dependencies
   - Run linters and code quality checks
   - Build application artifacts
   - Run unit tests

2. **Test Phase**
   - Run integration tests
   - Run E2E tests
   - Generate test coverage reports

3. **Deploy Phase**
   - Deploy to target environment
   - Run smoke tests
   - Update deployment status

4. **Notify Phase**
   - Send notifications to team channels
   - Update issue tracking systems
   - Generate deployment reports

## Atlassian Integration

### Jira Integration

Connect your deployment pipeline with Jira for automated issue tracking and project management.

#### Setup Steps

1. **Generate Jira API Token**
   - Go to [Atlassian Account Settings](https://id.atlassian.com/manage-profile/security/api-tokens)
   - Create a new API token
   - Store securely in GitHub Secrets as `JIRA_API_TOKEN`

2. **Configure Jira Connection**
   
   ```javascript
   // Add to your deployment scripts
   const jiraClient = {
     host: process.env.JIRA_HOST,
     username: process.env.JIRA_USERNAME,
     password: process.env.JIRA_API_TOKEN
   };
   ```

3. **Automate Issue Updates**
   
   ```yaml
   # In GitHub Actions workflow
   - name: Update Jira Issue
     uses: atlassian/gajira-transition@v3
     with:
       issue: ${{ github.event.issue.key }}
       transition: "Deploy Complete"
   ```

#### Automation Capabilities

- Auto-transition issues on deployment
- Link commits to Jira issues
- Create deployment tickets automatically
- Sync sprint planning with releases

### Confluence Integration

Document your deployments and maintain a knowledge base.

#### Setup Steps

1. **Generate Confluence API Token**
   - Use the same API token from Jira setup
   - Store in GitHub Secrets as `CONFLUENCE_API_TOKEN`

2. **Auto-Update Documentation**
   
   ```javascript
   // Example: Update deployment page
   const confluenceClient = {
     host: process.env.CONFLUENCE_HOST,
     username: process.env.CONFLUENCE_USERNAME,
     password: process.env.CONFLUENCE_API_TOKEN
   };
   
   // Update page with deployment info
   await confluenceClient.updatePage({
     id: pageId,
     version: currentVersion + 1,
     title: 'Deployment History',
     body: deploymentReport
   });
   ```

## Slack Integration

### Slack Bot Setup

1. **Create Slack App**
   - Go to [Slack API](https://api.slack.com/apps)
   - Create new app
   - Add Bot Token Scopes: `chat:write`, `channels:read`, `commands`

2. **Install to Workspace**
   - Install app to your workspace
   - Copy Bot Token to GitHub Secrets as `SLACK_BOT_TOKEN`

3. **Configure Webhook**
   - Create Incoming Webhook
   - Copy webhook URL to GitHub Secrets as `SLACK_WEBHOOK_URL`

### Deployment Notifications

```yaml
# In GitHub Actions workflow
- name: Notify Slack
  uses: slackapi/slack-github-action@v1
  with:
    channel-id: 'deployments'
    slack-message: |
      🚀 *Deployment Started*
      Environment: ${{ github.event.inputs.environment }}
      Version: ${{ github.sha }}
      Triggered by: ${{ github.actor }}
  env:
    SLACK_BOT_TOKEN: ${{ secrets.SLACK_BOT_TOKEN }}
```

### Slack Commands

Implement slash commands for deployment operations:

```
/deploy production       - Deploy to production
/deploy staging          - Deploy to staging
/status                  - Check deployment status
/rollback [version]      - Rollback to previous version
```

### Channel Setup

Recommended channel structure:
- `#deployments` - All deployment notifications
- `#build-status` - Build and test results
- `#incidents` - Production issues and alerts
- `#releases` - Release announcements

## Security Best Practices

### Secret Management

1. **Use GitHub Secrets**
   - Never commit secrets to repository
   - Use environment-specific secrets
   - Rotate secrets regularly

2. **Least Privilege Access**
   - Grant minimum required permissions
   - Use service accounts for automation
   - Review access logs regularly

3. **Encryption**
   - Enable encryption at rest
   - Use HTTPS for all communications
   - Implement certificate pinning

### Compliance

- **Audit Logging**: Enable comprehensive audit logs
- **Access Controls**: Implement RBAC (Role-Based Access Control)
- **Data Privacy**: Comply with GDPR, HIPAA, or other regulations
- **Vulnerability Scanning**: Regular security scans

## Monitoring and Observability

### Application Monitoring

1. **Azure Application Insights**
   
   ```javascript
   // Add to your application
   const appInsights = require('applicationinsights');
   appInsights.setup(process.env.APPLICATIONINSIGHTS_CONNECTION_STRING)
     .setAutoCollectRequests(true)
     .setAutoCollectPerformance(true)
     .setAutoCollectExceptions(true)
     .start();
   ```

2. **Custom Metrics**
   
   Track deployment-specific metrics:
   - Deployment frequency
   - Lead time for changes
   - Mean time to recovery (MTTR)
   - Change failure rate

### Health Checks

Implement health check endpoints:

```javascript
app.get('/health', (req, res) => {
  res.json({
    status: 'healthy',
    version: process.env.APP_VERSION,
    environment: process.env.ENVIRONMENT,
    uptime: process.uptime()
  });
});
```

### Alerting

Configure alerts for:
- Deployment failures
- High error rates
- Performance degradation
- Security incidents

## Troubleshooting

### Common Issues

#### Deployment Fails

1. Check GitHub Actions logs
2. Verify all secrets are configured
3. Ensure Azure resources exist
4. Check network connectivity

#### Integration Issues

1. Verify API tokens are valid
2. Check API rate limits
3. Review webhook configurations
4. Test connectivity manually

### Support Resources

- [Microsoft 365 Agents Toolkit Documentation](https://aka.ms/teamsfx-docs)
- [GitHub Actions Documentation](https://docs.github.com/actions)
- [Azure DevOps Documentation](https://docs.microsoft.com/azure/devops)
- [Atlassian Developer Documentation](https://developer.atlassian.com/)
- [Slack API Documentation](https://api.slack.com/)

## Additional Resources

- [CI/CD Best Practices](https://docs.microsoft.com/azure/devops/learn/what-is-continuous-integration)
- [Azure Security Best Practices](https://docs.microsoft.com/azure/security/fundamentals/best-practices-and-patterns)
- [GitHub Enterprise Documentation](https://docs.github.com/enterprise)

## Conclusion

This enterprise deployment guide provides a foundation for deploying Microsoft 365 Agents in production environments. Customize the workflows and integrations based on your organization's specific requirements and security policies.

For questions or support, please refer to the [support documentation](../SUPPORT.md) or reach out to the Microsoft 365 Agents Toolkit team.
