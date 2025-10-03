# GitHub Enterprise Integration

This guide covers integrating Microsoft 365 Agents Toolkit with GitHub Enterprise for enterprise-scale deployments.

## Overview

GitHub Enterprise provides additional features for large organizations including:
- Advanced security features
- SAML single sign-on (SSO)
- Audit logging
- Private networking
- Enterprise-grade support

## Prerequisites

- GitHub Enterprise account with admin access
- Microsoft 365 Agents Toolkit installed
- Azure subscription (for cloud deployments)

## Setup GitHub Enterprise Repository

### 1. Create Repository

```bash
# Create a new repository in your GitHub Enterprise instance
gh repo create your-org/m365-agents-app --private --description "Microsoft 365 Agents Application"

# Clone the repository
git clone https://github.yourcompany.com/your-org/m365-agents-app.git
cd m365-agents-app
```

### 2. Initialize M365 Agents Project

```bash
# Install M365 Agents Toolkit CLI
npm install -g @microsoft/m365agentstoolkit-cli

# Create a new project
m365agents new
# Follow the prompts to select your template and configuration
```

### 3. Configure Repository Settings

#### Enable Branch Protection

1. Go to repository Settings > Branches
2. Add branch protection rule for `main`:
   - Require pull request reviews before merging
   - Require status checks to pass before merging
   - Require signed commits
   - Include administrators

#### Set Up Environments

1. Go to Settings > Environments
2. Create environments:
   - `development`
   - `staging`
   - `production`
3. Configure protection rules:
   - Required reviewers for production
   - Wait timer for staging (optional)
   - Deployment branches (limit to specific branches)

## Configure GitHub Actions

### 1. Add Repository Secrets

Navigate to Settings > Secrets and variables > Actions, then add:

```
Required Secrets:
- BOT_ID                    # Microsoft Bot Framework App ID
- BOT_PASSWORD              # Microsoft Bot Framework App Password
- AZURE_CREDENTIALS         # Azure service principal credentials
- AZURE_SUBSCRIPTION_ID     # Azure subscription ID
- AZURE_RESOURCE_GROUP      # Azure resource group name
- TEAMS_APP_ID              # Microsoft Teams App ID
- TEAMS_APP_PASSWORD        # Teams App Password
```

For Azure credentials, create a service principal:

```bash
az ad sp create-for-rbac --name "m365-agents-github-actions" \
  --role contributor \
  --scopes /subscriptions/{subscription-id}/resourceGroups/{resource-group} \
  --sdk-auth
```

Store the JSON output as `AZURE_CREDENTIALS` secret.

### 2. Create Workflow

Copy the example workflow to your repository:

```bash
cp .github/workflows/enterprise-deploy.yml.example .github/workflows/deploy.yml
```

Customize the workflow based on your requirements.

### 3. Configure GitHub Actions Permissions

Set Actions permissions in Settings > Actions > General:
- Allow all actions and reusable workflows
- Set Workflow permissions to "Read and write permissions"
- Enable "Allow GitHub Actions to create and approve pull requests"

## Security Configuration

### Enable SAML SSO

If using SAML authentication:

1. Go to Organization Settings > Security > SAML single sign-on
2. Enable SAML SSO
3. Configure with your identity provider
4. Require SSO for all members

### Enable Audit Logging

1. Go to Organization Settings > Audit log
2. Enable audit log streaming (for real-time monitoring)
3. Configure destination (e.g., Splunk, Azure Monitor)

### Security Scanning

Enable GitHub security features:

```yaml
# .github/workflows/security.yml
name: Security Scanning

on:
  push:
    branches: [ main ]
  pull_request:
    branches: [ main ]
  schedule:
    - cron: '0 0 * * 0'  # Weekly scan

jobs:
  security:
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v4
      
      - name: Run CodeQL Analysis
        uses: github/codeql-action/init@v3
        with:
          languages: 'javascript, typescript'
      
      - name: Perform CodeQL Analysis
        uses: github/codeql-action/analyze@v3
      
      - name: Run Dependency Review
        uses: actions/dependency-review-action@v4
```

## Advanced Features

### GitHub Container Registry

Use GitHub Container Registry for Docker images:

```yaml
# In your workflow
- name: Build and push Docker image
  run: |
    echo ${{ secrets.GITHUB_TOKEN }} | docker login ghcr.io -u ${{ github.actor }} --password-stdin
    docker build -t ghcr.io/${{ github.repository }}/app:${{ github.sha }} .
    docker push ghcr.io/${{ github.repository }}/app:${{ github.sha }}
```

### GitHub Packages

Store npm packages in GitHub Packages:

```yaml
# .npmrc
@your-org:registry=https://npm.pkg.github.com
//npm.pkg.github.com/:_authToken=${NODE_AUTH_TOKEN}
```

```yaml
# In your workflow
- name: Publish to GitHub Packages
  run: npm publish
  env:
    NODE_AUTH_TOKEN: ${{ secrets.GITHUB_TOKEN }}
```

### Self-Hosted Runners

For sensitive workloads, use self-hosted runners:

1. Go to Settings > Actions > Runners
2. Click "New self-hosted runner"
3. Follow setup instructions for your OS
4. Update workflows to use self-hosted runner:

```yaml
jobs:
  deploy:
    runs-on: self-hosted  # Use self-hosted runner
```

## CI/CD Best Practices

### 1. Workflow Organization

```
.github/
  workflows/
    ci.yml              # Continuous Integration
    cd-dev.yml          # Deploy to Development
    cd-staging.yml      # Deploy to Staging
    cd-production.yml   # Deploy to Production
    security.yml        # Security Scanning
    cleanup.yml         # Cleanup old deployments
```

### 2. Reusable Workflows

Create reusable workflows for common tasks:

```yaml
# .github/workflows/deploy-template.yml
name: Deploy Template

on:
  workflow_call:
    inputs:
      environment:
        required: true
        type: string
    secrets:
      bot_id:
        required: true
      bot_password:
        required: true

jobs:
  deploy:
    runs-on: ubuntu-latest
    environment: ${{ inputs.environment }}
    steps:
      - uses: actions/checkout@v4
      # Add deployment steps
```

Use in other workflows:

```yaml
jobs:
  deploy-dev:
    uses: ./.github/workflows/deploy-template.yml
    with:
      environment: development
    secrets:
      bot_id: ${{ secrets.BOT_ID }}
      bot_password: ${{ secrets.BOT_PASSWORD }}
```

### 3. Matrix Builds

Test across multiple configurations:

```yaml
jobs:
  test:
    runs-on: ubuntu-latest
    strategy:
      matrix:
        node-version: [18, 20, 22]
    steps:
      - uses: actions/setup-node@v4
        with:
          node-version: ${{ matrix.node-version }}
```

## Monitoring and Observability

### GitHub Actions Insights

Track workflow performance:
1. Go to Actions tab
2. View workflow runs
3. Analyze run duration, success rates
4. Identify bottlenecks

### Workflow Notifications

Configure notifications:

```yaml
# Send notifications on workflow failure
- name: Notify on Failure
  if: failure()
  uses: actions/github-script@v7
  with:
    script: |
      github.rest.issues.create({
        owner: context.repo.owner,
        repo: context.repo.repo,
        title: 'Deployment Failed',
        body: 'Workflow run failed: ${{ github.server_url }}/${{ github.repository }}/actions/runs/${{ github.run_id }}'
      })
```

## Troubleshooting

### Common Issues

#### Authentication Errors

If GitHub Actions cannot authenticate with Azure:

1. Verify `AZURE_CREDENTIALS` secret is valid
2. Check service principal permissions
3. Ensure subscription ID is correct

#### Workflow Not Triggering

1. Check branch protection rules
2. Verify workflow triggers in YAML
3. Check Actions permissions

#### Secret Not Found

1. Verify secret name matches exactly (case-sensitive)
2. Check secret is available in the environment
3. Ensure secret scope is correct (repository/organization/environment)

### Debug Mode

Enable debug logging:

1. Go to Settings > Secrets
2. Add `ACTIONS_STEP_DEBUG` with value `true`
3. Re-run workflow to see detailed logs

## Resources

- [GitHub Enterprise Documentation](https://docs.github.com/enterprise)
- [GitHub Actions Documentation](https://docs.github.com/actions)
- [Microsoft 365 Agents Toolkit](https://aka.ms/teamsfx-docs)
- [Azure DevOps Integration](https://docs.microsoft.com/azure/devops)

## Next Steps

- [Set up Atlassian Integration](./atlassian-integration.md)
- [Configure Slack Integration](./slack-integration.md)
- [Review Security Best Practices](../enterprise-deployment.md#security-best-practices)
