# Enterprise Integrations

This directory contains documentation for integrating Microsoft 365 Agents Toolkit with enterprise tools and platforms.

## Available Integrations

### [GitHub Enterprise Integration](./github-enterprise.md)
Complete guide for setting up CI/CD with GitHub Enterprise, including:
- Repository setup and configuration
- GitHub Actions workflows
- Security scanning and compliance
- Self-hosted runners
- Container registry

### [Atlassian Integration](./atlassian-integration.md)
Integration with Jira and Confluence for project management and documentation:
- Jira issue tracking and automation
- Smart commits
- Confluence documentation updates
- Release management
- Sprint integration

### [Slack Integration](./slack-integration.md)
Real-time notifications and ChatOps with Slack:
- Deployment notifications
- Interactive messages and commands
- Channel organization
- Incident alerts
- Metrics and reporting

## Quick Start

1. **Choose Your Integrations**: Select the tools your team uses
2. **Set Up Credentials**: Follow setup guides to create API tokens
3. **Configure GitHub Secrets**: Store credentials securely
4. **Implement Workflows**: Use provided examples as templates
5. **Test**: Verify integrations work correctly
6. **Customize**: Adapt to your team's needs

## Integration Comparison

| Feature | GitHub Enterprise | Atlassian | Slack |
|---------|------------------|-----------|-------|
| **Primary Use** | CI/CD & Code Management | Project Tracking | Communication |
| **Automation** | ✅ Workflows | ✅ Rules | ✅ Bots |
| **Notifications** | ✅ Email/Web | ✅ Email | ✅ Real-time |
| **API Available** | ✅ REST & GraphQL | ✅ REST | ✅ REST & WebSocket |
| **Cost** | Enterprise | Cloud/Server | Free/Paid |

## Best Practices

### Security
- Store all credentials in GitHub Secrets
- Rotate API tokens regularly
- Use least privilege access
- Enable audit logging
- Review permissions quarterly

### Automation
- Start with simple notifications
- Add interactivity gradually
- Monitor rate limits
- Implement error handling
- Test in non-production first

### Team Collaboration
- Establish naming conventions
- Document custom workflows
- Train team on new tools
- Gather feedback regularly
- Iterate on processes

## Common Workflows

### Deployment Pipeline
1. **Code Push** → GitHub Enterprise triggers workflow
2. **Build & Test** → Results posted to Slack
3. **Deploy** → Update Jira issue status
4. **Success** → Update Confluence docs, notify Slack

### Issue Tracking
1. **Create Jira Issue** → Link to GitHub PR
2. **PR Merged** → Update Jira status
3. **Deploy** → Add deployment comment to Jira
4. **Verification** → Close Jira issue

### Incident Response
1. **Failure Detected** → Alert Slack channel
2. **Create Jira Incident** → Auto-assign to on-call
3. **Fix Deployed** → Update all systems
4. **Post-Mortem** → Document in Confluence

## Troubleshooting

### Authentication Issues
- Verify tokens are valid and not expired
- Check token permissions/scopes
- Ensure secrets are correctly named
- Test authentication separately

### Notification Failures
- Check API rate limits
- Verify webhook URLs are correct
- Review error logs in GitHub Actions
- Test with minimal examples first

### Integration Not Working
- Review documentation for each service
- Check service status pages
- Verify network connectivity
- Enable debug logging

## Support

For issues specific to:
- **GitHub Enterprise**: [GitHub Enterprise Support](https://enterprise.github.com/support)
- **Atlassian**: [Atlassian Support](https://support.atlassian.com/)
- **Slack**: [Slack Help Center](https://slack.com/help)
- **Microsoft 365 Agents Toolkit**: [Toolkit Support](../../SUPPORT.md)

## Additional Resources

- [Enterprise Deployment Guide](../enterprise-deployment.md)
- [Microsoft 365 Agents Toolkit Documentation](https://aka.ms/teamsfx-docs)
- [GitHub Actions Documentation](https://docs.github.com/actions)
- [Atlassian Developer Documentation](https://developer.atlassian.com/)
- [Slack API Documentation](https://api.slack.com/)

## Contributing

Have a new integration or improvement to existing ones? See our [Contributing Guide](../../CONTRIBUTING.md) for information on how to contribute.
