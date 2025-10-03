# Atlassian Integration (Jira & Confluence)

This guide covers integrating Microsoft 365 Agents Toolkit deployments with Atlassian products (Jira and Confluence).

## Overview

Integrate your deployment pipeline with Atlassian tools to:
- Automatically track deployment status in Jira
- Update issue statuses based on deployment events
- Maintain deployment documentation in Confluence
- Link code changes to Jira issues
- Generate release reports

## Prerequisites

- Atlassian Cloud account (Jira and/or Confluence)
- Admin access to create API tokens
- GitHub repository with Microsoft 365 Agents project

## Jira Integration

### Setup

#### 1. Generate Jira API Token

1. Log in to [Atlassian Account](https://id.atlassian.com/)
2. Go to Security > Create and manage API tokens
3. Click "Create API token"
4. Give it a name (e.g., "GitHub Actions")
5. Copy the token immediately (you won't see it again)

#### 2. Store Credentials in GitHub

Add these secrets to your GitHub repository:

```
JIRA_BASE_URL       # e.g., https://yourcompany.atlassian.net
JIRA_USER_EMAIL     # Your Atlassian account email
JIRA_API_TOKEN      # Token from step 1
JIRA_PROJECT_KEY    # Your Jira project key (e.g., "PROJ")
```

#### 3. Configure Jira Actions in Workflow

```yaml
# .github/workflows/deploy.yml
name: Deploy with Jira Integration

on:
  push:
    branches: [ main ]

jobs:
  deploy:
    runs-on: ubuntu-latest
    steps:
      - name: Checkout
        uses: actions/checkout@v4
      
      - name: Extract Jira Issue Key
        id: jira
        run: |
          # Extract issue key from commit message or branch name
          ISSUE_KEY=$(git log -1 --pretty=%B | grep -o '[A-Z]\+-[0-9]\+' | head -1)
          echo "issue_key=$ISSUE_KEY" >> $GITHUB_OUTPUT
      
      - name: Transition Jira Issue
        if: steps.jira.outputs.issue_key
        uses: atlassian/gajira-transition@v3
        with:
          issue: ${{ steps.jira.outputs.issue_key }}
          transition: "In Progress"
        env:
          JIRA_BASE_URL: ${{ secrets.JIRA_BASE_URL }}
          JIRA_USER_EMAIL: ${{ secrets.JIRA_USER_EMAIL }}
          JIRA_API_TOKEN: ${{ secrets.JIRA_API_TOKEN }}
      
      # Your deployment steps here
      
      - name: Update Jira on Success
        if: success() && steps.jira.outputs.issue_key
        uses: atlassian/gajira-transition@v3
        with:
          issue: ${{ steps.jira.outputs.issue_key }}
          transition: "Deployed"
        env:
          JIRA_BASE_URL: ${{ secrets.JIRA_BASE_URL }}
          JIRA_USER_EMAIL: ${{ secrets.JIRA_USER_EMAIL }}
          JIRA_API_TOKEN: ${{ secrets.JIRA_API_TOKEN }}
      
      - name: Comment on Jira Issue
        if: success() && steps.jira.outputs.issue_key
        uses: atlassian/gajira-comment@v3
        with:
          issue: ${{ steps.jira.outputs.issue_key }}
          comment: |
            ✅ Deployment completed successfully
            
            Environment: Production
            Version: ${{ github.sha }}
            Workflow: ${{ github.server_url }}/${{ github.repository }}/actions/runs/${{ github.run_id }}
        env:
          JIRA_BASE_URL: ${{ secrets.JIRA_BASE_URL }}
          JIRA_USER_EMAIL: ${{ secrets.JIRA_USER_EMAIL }}
          JIRA_API_TOKEN: ${{ secrets.JIRA_API_TOKEN }}
```

### Jira Automation Rules

Set up Jira automation to respond to deployments:

#### 1. Create Deployment Status Field

1. Go to Project Settings > Issue Types > Select type
2. Add custom field: "Deployment Status"
3. Set options: Not Deployed, In Progress, Deployed, Failed

#### 2. Create Automation Rule

1. Go to Project Settings > Automation
2. Create new rule with trigger: "Field value changed"
3. Select "Deployment Status" field
4. Add actions:
   - Send notification to team
   - Update sprint board
   - Trigger follow-up tasks

#### 3. Smart Commits

Use smart commits to automate Jira updates:

```bash
# Commit format: <ISSUE_KEY> #comment <message> #time <time>
git commit -m "PROJ-123 #comment Fixed deployment issue #time 2h"

# Transition issue
git commit -m "PROJ-123 #done Deployment completed"

# Multiple actions
git commit -m "PROJ-123 #comment Updated config #time 1h #transition In Review"
```

### Jira REST API Integration

For custom integrations, use the Jira REST API:

```javascript
// update-jira.js
const axios = require('axios');

async function updateJiraIssue(issueKey, status, comment) {
  const auth = Buffer.from(
    `${process.env.JIRA_USER_EMAIL}:${process.env.JIRA_API_TOKEN}`
  ).toString('base64');

  const baseUrl = process.env.JIRA_BASE_URL;
  
  // Add comment
  await axios.post(
    `${baseUrl}/rest/api/3/issue/${issueKey}/comment`,
    {
      body: {
        type: 'doc',
        version: 1,
        content: [
          {
            type: 'paragraph',
            content: [{ type: 'text', text: comment }]
          }
        ]
      }
    },
    {
      headers: {
        'Authorization': `Basic ${auth}`,
        'Content-Type': 'application/json'
      }
    }
  );
  
  // Transition issue
  const transitions = await axios.get(
    `${baseUrl}/rest/api/3/issue/${issueKey}/transitions`,
    {
      headers: { 'Authorization': `Basic ${auth}` }
    }
  );
  
  const transition = transitions.data.transitions.find(
    t => t.name === status
  );
  
  if (transition) {
    await axios.post(
      `${baseUrl}/rest/api/3/issue/${issueKey}/transitions`,
      { transition: { id: transition.id } },
      {
        headers: {
          'Authorization': `Basic ${auth}`,
          'Content-Type': 'application/json'
        }
      }
    );
  }
}

// Usage in GitHub Actions
const issueKey = process.argv[2];
const status = process.argv[3];
const comment = process.argv[4];

updateJiraIssue(issueKey, status, comment)
  .then(() => console.log('Jira updated successfully'))
  .catch(err => console.error('Error updating Jira:', err));
```

## Confluence Integration

### Setup

#### 1. Get Confluence Credentials

Use the same API token from Jira setup. Add to GitHub secrets:

```
CONFLUENCE_BASE_URL    # e.g., https://yourcompany.atlassian.net/wiki
CONFLUENCE_SPACE_KEY   # Your Confluence space key
```

#### 2. Create Deployment Documentation Page

Create a template page in Confluence:

```markdown
# Deployment History

## Latest Deployment

**Date:** 
**Version:** 
**Environment:** 
**Status:** 
**Deployed by:** 

## Deployment Log

| Date | Version | Environment | Status | Deployed By |
|------|---------|-------------|--------|-------------|
|      |         |             |        |             |
```

Get the page ID from the URL when editing the page.

#### 3. Update Confluence in Workflow

```yaml
- name: Update Confluence Documentation
  run: |
    node update-confluence.js \
      "${{ secrets.CONFLUENCE_PAGE_ID }}" \
      "${{ github.sha }}" \
      "production" \
      "${{ job.status }}" \
      "${{ github.actor }}"
  env:
    CONFLUENCE_BASE_URL: ${{ secrets.CONFLUENCE_BASE_URL }}
    CONFLUENCE_USER_EMAIL: ${{ secrets.JIRA_USER_EMAIL }}
    CONFLUENCE_API_TOKEN: ${{ secrets.JIRA_API_TOKEN }}
```

### Confluence API Script

```javascript
// update-confluence.js
const axios = require('axios');

async function updateConfluencePage(pageId, version, environment, status, deployedBy) {
  const auth = Buffer.from(
    `${process.env.CONFLUENCE_USER_EMAIL}:${process.env.CONFLUENCE_API_TOKEN}`
  ).toString('base64');

  const baseUrl = process.env.CONFLUENCE_BASE_URL;
  
  // Get current page
  const page = await axios.get(
    `${baseUrl}/rest/api/content/${pageId}?expand=body.storage,version`,
    {
      headers: { 'Authorization': `Basic ${auth}` }
    }
  );
  
  const currentContent = page.data.body.storage.value;
  const currentVersion = page.data.version.number;
  
  // Update content
  const now = new Date().toISOString();
  const newRow = `
    <tr>
      <td>${now}</td>
      <td>${version}</td>
      <td>${environment}</td>
      <td>${status}</td>
      <td>${deployedBy}</td>
    </tr>
  `;
  
  const updatedContent = currentContent.replace(
    '</tbody>',
    newRow + '</tbody>'
  );
  
  // Update page
  await axios.put(
    `${baseUrl}/rest/api/content/${pageId}`,
    {
      id: pageId,
      type: 'page',
      title: page.data.title,
      version: { number: currentVersion + 1 },
      body: {
        storage: {
          value: updatedContent,
          representation: 'storage'
        }
      }
    },
    {
      headers: {
        'Authorization': `Basic ${auth}`,
        'Content-Type': 'application/json'
      }
    }
  );
  
  console.log('Confluence page updated successfully');
}

// Usage
const [pageId, version, environment, status, deployedBy] = process.argv.slice(2);
updateConfluencePage(pageId, version, environment, status, deployedBy)
  .catch(err => console.error('Error:', err));
```

### Automated Documentation

Create comprehensive documentation automatically:

```yaml
- name: Generate and Upload Documentation
  run: |
    # Generate deployment report
    cat << EOF > deployment-report.html
    <h2>Deployment Report - $(date)</h2>
    <h3>Summary</h3>
    <ul>
      <li>Version: ${{ github.sha }}</li>
      <li>Environment: Production</li>
      <li>Deployed by: ${{ github.actor }}</li>
      <li>Status: Success</li>
    </ul>
    <h3>Changes</h3>
    $(git log --oneline -10 | sed 's/^/<li>/' | sed 's/$/<\/li>/')
    EOF
    
    node upload-to-confluence.js deployment-report.html
```

## Advanced Integration Patterns

### 1. Release Management

Track releases across Jira and Confluence:

```yaml
- name: Create Jira Release
  run: |
    curl -X POST \
      -H "Authorization: Basic $(echo -n $JIRA_USER_EMAIL:$JIRA_API_TOKEN | base64)" \
      -H "Content-Type: application/json" \
      -d '{
        "name": "v${{ github.run_number }}",
        "description": "Production release",
        "released": true,
        "releaseDate": "'$(date -I)'"
      }' \
      "$JIRA_BASE_URL/rest/api/3/version"
```

### 2. Sprint Integration

Link deployments to sprints:

```javascript
// Get active sprint
const sprint = await axios.get(
  `${jiraBaseUrl}/rest/agile/1.0/board/${boardId}/sprint?state=active`
);

// Add deployment info to sprint
await axios.post(
  `${jiraBaseUrl}/rest/api/3/issue/${issueKey}`,
  {
    fields: {
      customfield_deployment: {
        version: version,
        date: new Date().toISOString(),
        environment: 'production'
      }
    }
  }
);
```

### 3. Metrics Dashboard

Create a Confluence dashboard with deployment metrics:

```markdown
## Deployment Metrics

**This Month:**
- Total Deployments: 42
- Success Rate: 95%
- Average Lead Time: 2.5 hours
- Mean Time to Recovery: 15 minutes

**Trends:**
[Embed chart/graph]
```

## Best Practices

1. **Consistent Naming**: Use consistent issue key format (e.g., PROJ-123)
2. **Automation**: Automate status updates to reduce manual work
3. **Documentation**: Keep Confluence docs updated automatically
4. **Notifications**: Set up alerts for deployment failures
5. **Archiving**: Archive old deployment records regularly

## Troubleshooting

### API Token Issues

If authentication fails:
1. Regenerate API token
2. Verify email is correct
3. Check token has required permissions

### Rate Limiting

Atlassian APIs have rate limits:
- Cloud: ~100 requests per minute
- Use exponential backoff for retries
- Cache responses when possible

### Confluence Page Not Updating

1. Verify page ID is correct
2. Check version number increments
3. Ensure content format is valid XHTML

## Resources

- [Jira REST API Documentation](https://developer.atlassian.com/cloud/jira/platform/rest/v3/)
- [Confluence REST API Documentation](https://developer.atlassian.com/cloud/confluence/rest/v2/)
- [Atlassian GitHub Actions](https://github.com/marketplace?query=atlassian)

## Next Steps

- [Configure Slack Integration](./slack-integration.md)
- [Review Enterprise Deployment Guide](../enterprise-deployment.md)
