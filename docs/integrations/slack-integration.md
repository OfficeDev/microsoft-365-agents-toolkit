# Slack Integration

This guide covers integrating Microsoft 365 Agents Toolkit deployments with Slack for real-time notifications and team collaboration.

## Overview

Integrate Slack with your deployment pipeline to:
- Send real-time deployment notifications
- Create interactive deployment commands
- Alert team on build failures
- Track deployment metrics in channels
- Enable ChatOps workflows

## Prerequisites

- Slack workspace with admin access
- GitHub repository with Microsoft 365 Agents project
- Ability to create and install Slack apps

## Quick Start

### 1. Create Slack App

1. Go to [Slack API Apps](https://api.slack.com/apps)
2. Click "Create New App"
3. Choose "From scratch"
4. Enter app name (e.g., "M365 Agents Deployer")
5. Select your workspace
6. Click "Create App"

### 2. Configure Bot Permissions

1. Go to "OAuth & Permissions" in the sidebar
2. Add these Bot Token Scopes:
   - `chat:write` - Send messages
   - `chat:write.public` - Send to channels without joining
   - `channels:read` - List public channels
   - `groups:read` - List private channels
   - `im:write` - Send direct messages
   - `commands` - Create slash commands
   - `files:write` - Upload files
   - `reactions:write` - Add emoji reactions

3. Install app to workspace
4. Copy the "Bot User OAuth Token" (starts with `xoxb-`)

### 3. Create Webhook (Alternative Method)

For simple notifications without a full app:

1. Go to your app settings
2. Navigate to "Incoming Webhooks"
3. Activate Incoming Webhooks
4. Click "Add New Webhook to Workspace"
5. Select channel and authorize
6. Copy the webhook URL

### 4. Store Credentials in GitHub

Add these secrets to your GitHub repository:

```
SLACK_BOT_TOKEN      # Bot token from step 2 (xoxb-...)
SLACK_WEBHOOK_URL    # Webhook URL from step 3
SLACK_CHANNEL_ID     # Channel ID for notifications (e.g., C0123456789)
```

To get channel ID:
- Right-click channel > View channel details
- Scroll down to find the Channel ID

## Basic Integration

### Webhook Notifications

Simple notifications using incoming webhooks:

```yaml
# .github/workflows/deploy.yml
name: Deploy with Slack Notifications

on:
  push:
    branches: [ main ]

jobs:
  deploy:
    runs-on: ubuntu-latest
    steps:
      - name: Notify Deployment Started
        run: |
          curl -X POST -H 'Content-type: application/json' \
            --data '{
              "text": "🚀 Deployment Started",
              "blocks": [
                {
                  "type": "section",
                  "text": {
                    "type": "mrkdwn",
                    "text": "*Deployment Started*\n\n• Repository: `${{ github.repository }}`\n• Branch: `${{ github.ref_name }}`\n• Commit: `${{ github.sha }}`\n• Triggered by: @${{ github.actor }}"
                  }
                }
              ]
            }' \
            ${{ secrets.SLACK_WEBHOOK_URL }}
      
      # Your deployment steps here
      
      - name: Notify Deployment Success
        if: success()
        run: |
          curl -X POST -H 'Content-type: application/json' \
            --data '{
              "text": "✅ Deployment Successful",
              "blocks": [
                {
                  "type": "section",
                  "text": {
                    "type": "mrkdwn",
                    "text": "✅ *Deployment Successful*\n\nEnvironment: Production\nVersion: `${{ github.sha }}`"
                  }
                }
              ]
            }' \
            ${{ secrets.SLACK_WEBHOOK_URL }}
      
      - name: Notify Deployment Failure
        if: failure()
        run: |
          curl -X POST -H 'Content-type: application/json' \
            --data '{
              "text": "❌ Deployment Failed",
              "blocks": [
                {
                  "type": "section",
                  "text": {
                    "type": "mrkdwn",
                    "text": "❌ *Deployment Failed*\n\nPlease check <${{ github.server_url }}/${{ github.repository }}/actions/runs/${{ github.run_id }}|the logs>"
                  }
                }
              ]
            }' \
            ${{ secrets.SLACK_WEBHOOK_URL }}
```

### Using Slack GitHub Action

More advanced notifications with the official action:

```yaml
- name: Send Slack Notification
  uses: slackapi/slack-github-action@v1.26.0
  with:
    channel-id: ${{ secrets.SLACK_CHANNEL_ID }}
    payload: |
      {
        "text": "Deployment Status: ${{ job.status }}",
        "blocks": [
          {
            "type": "header",
            "text": {
              "type": "plain_text",
              "text": "🚀 M365 Agents Deployment"
            }
          },
          {
            "type": "section",
            "fields": [
              {
                "type": "mrkdwn",
                "text": "*Status:*\n${{ job.status }}"
              },
              {
                "type": "mrkdwn",
                "text": "*Environment:*\nProduction"
              },
              {
                "type": "mrkdwn",
                "text": "*Version:*\n${{ github.sha }}"
              },
              {
                "type": "mrkdwn",
                "text": "*Triggered By:*\n@${{ github.actor }}"
              }
            ]
          },
          {
            "type": "actions",
            "elements": [
              {
                "type": "button",
                "text": {
                  "type": "plain_text",
                  "text": "View Workflow"
                },
                "url": "${{ github.server_url }}/${{ github.repository }}/actions/runs/${{ github.run_id }}"
              },
              {
                "type": "button",
                "text": {
                  "type": "plain_text",
                  "text": "View Commit"
                },
                "url": "${{ github.server_url }}/${{ github.repository }}/commit/${{ github.sha }}"
              }
            ]
          }
        ]
      }
  env:
    SLACK_BOT_TOKEN: ${{ secrets.SLACK_BOT_TOKEN }}
```

## Advanced Features

### Interactive Messages

Create messages with buttons and interactive elements:

```yaml
- name: Send Interactive Deployment Approval
  uses: slackapi/slack-github-action@v1.26.0
  with:
    channel-id: ${{ secrets.SLACK_CHANNEL_ID }}
    payload: |
      {
        "text": "Deployment Approval Required",
        "blocks": [
          {
            "type": "section",
            "text": {
              "type": "mrkdwn",
              "text": "⚠️ *Production Deployment Requires Approval*\n\nVersion: ${{ github.sha }}\nChanges: ${{ github.event.head_commit.message }}"
            }
          },
          {
            "type": "actions",
            "elements": [
              {
                "type": "button",
                "text": {
                  "type": "plain_text",
                  "text": "✅ Approve"
                },
                "style": "primary",
                "action_id": "approve_deployment"
              },
              {
                "type": "button",
                "text": {
                  "type": "plain_text",
                  "text": "❌ Reject"
                },
                "style": "danger",
                "action_id": "reject_deployment"
              }
            ]
          }
        ]
      }
  env:
    SLACK_BOT_TOKEN: ${{ secrets.SLACK_BOT_TOKEN }}
```

### Thread Replies

Keep deployment updates in threads:

```javascript
// send-slack-message.js
const { WebClient } = require('@slack/web-api');

const slack = new WebClient(process.env.SLACK_BOT_TOKEN);

async function sendMessage(channel, text, threadTs = null) {
  const result = await slack.chat.postMessage({
    channel: channel,
    text: text,
    thread_ts: threadTs  // Reply to thread if provided
  });
  return result.ts;  // Return message timestamp for threading
}

// Start deployment thread
const mainTs = await sendMessage('C0123456789', '🚀 Deployment started');

// Add updates as thread replies
await sendMessage('C0123456789', '⏳ Building...', mainTs);
await sendMessage('C0123456789', '⏳ Testing...', mainTs);
await sendMessage('C0123456789', '✅ Deployment complete!', mainTs);
```

### Slash Commands

Create slash commands for deployment operations:

#### 1. Register Slash Command

1. Go to your Slack app settings
2. Navigate to "Slash Commands"
3. Click "Create New Command"
4. Enter command details:
   - Command: `/deploy`
   - Request URL: `https://your-api.com/slack/commands`
   - Short Description: "Deploy M365 Agents"
   - Usage Hint: `[environment] [version]`

#### 2. Create Command Handler

```javascript
// slack-commands.js
const express = require('express');
const { WebClient } = require('@slack/web-api');

const app = express();
app.use(express.urlencoded({ extended: true }));

const slack = new WebClient(process.env.SLACK_BOT_TOKEN);

app.post('/slack/commands', async (req, res) => {
  const { command, text, user_name, response_url } = req.body;
  
  if (command === '/deploy') {
    const [environment, version] = text.split(' ');
    
    // Acknowledge command immediately
    res.json({
      response_type: 'in_channel',
      text: `🚀 Deployment to ${environment} initiated by @${user_name}`
    });
    
    // Trigger GitHub Actions workflow
    // You'll need to implement this part using GitHub API
    
    // Send follow-up message
    await slack.chat.postMessage({
      channel: req.body.channel_id,
      text: `Deployment workflow started. Check progress: ${workflowUrl}`
    });
  }
});

app.listen(3000);
```

### Status Updates with Emoji Reactions

Add emoji reactions to track deployment progress:

```javascript
// update-message-status.js
const { WebClient } = require('@slack/web-api');

const slack = new WebClient(process.env.SLACK_BOT_TOKEN);

async function updateDeploymentStatus(channel, timestamp, status) {
  const emojiMap = {
    'started': 'rocket',
    'building': 'hammer',
    'testing': 'microscope',
    'deploying': 'package',
    'success': 'white_check_mark',
    'failed': 'x'
  };
  
  await slack.reactions.add({
    channel: channel,
    timestamp: timestamp,
    name: emojiMap[status]
  });
}
```

## Channel Organization

### Recommended Channel Structure

```
#deployments           - All deployment notifications
#deployments-dev       - Development environment
#deployments-staging   - Staging environment
#deployments-prod      - Production environment
#build-status          - CI/CD build results
#incidents             - Production incidents
#releases              - Release announcements
#deployment-approvals  - Approval requests
```

### Channel-Specific Notifications

```yaml
- name: Notify Appropriate Channel
  run: |
    if [ "${{ github.ref }}" == "refs/heads/main" ]; then
      CHANNEL="C0123456789"  # #deployments-prod
    elif [ "${{ github.ref }}" == "refs/heads/staging" ]; then
      CHANNEL="C9876543210"  # #deployments-staging
    else
      CHANNEL="C1234567890"  # #deployments-dev
    fi
    
    curl -X POST -H 'Content-type: application/json' \
      -H "Authorization: Bearer ${{ secrets.SLACK_BOT_TOKEN }}" \
      --data "{\"channel\":\"$CHANNEL\",\"text\":\"Deployment started\"}" \
      https://slack.com/api/chat.postMessage
```

## Metrics and Reporting

### Daily Deployment Summary

```yaml
# .github/workflows/daily-summary.yml
name: Daily Deployment Summary

on:
  schedule:
    - cron: '0 18 * * *'  # 6 PM UTC daily

jobs:
  summary:
    runs-on: ubuntu-latest
    steps:
      - name: Generate Summary
        run: |
          # Fetch deployment data from last 24 hours
          DEPLOYMENTS=$(gh api /repos/${{ github.repository }}/actions/runs \
            --jq '.workflow_runs[] | select(.created_at > (now - 86400 | todate))')
          
          # Create summary message
          SUMMARY="📊 *Daily Deployment Summary*\n\n"
          SUMMARY+="Total Deployments: $COUNT\n"
          SUMMARY+="Success Rate: $SUCCESS_RATE%\n"
          SUMMARY+="Average Duration: $AVG_DURATION"
      
      - name: Send to Slack
        uses: slackapi/slack-github-action@v1.26.0
        with:
          channel-id: ${{ secrets.SLACK_CHANNEL_ID }}
          payload: |
            {
              "text": "${{ env.SUMMARY }}"
            }
        env:
          SLACK_BOT_TOKEN: ${{ secrets.SLACK_BOT_TOKEN }}
```

### Incident Alerts

```yaml
- name: Alert on Failure
  if: failure()
  uses: slackapi/slack-github-action@v1.26.0
  with:
    channel-id: ${{ secrets.SLACK_INCIDENTS_CHANNEL }}
    payload: |
      {
        "text": "🚨 PRODUCTION INCIDENT",
        "blocks": [
          {
            "type": "header",
            "text": {
              "type": "plain_text",
              "text": "🚨 Production Deployment Failed"
            }
          },
          {
            "type": "section",
            "text": {
              "type": "mrkdwn",
              "text": "*Severity:* High\n*Environment:* Production\n*Time:* $(date -u +\"%Y-%m-%d %H:%M:%S UTC\")"
            }
          },
          {
            "type": "section",
            "text": {
              "type": "mrkdwn",
              "text": "<!channel> Immediate attention required!"
            }
          }
        ]
      }
  env:
    SLACK_BOT_TOKEN: ${{ secrets.SLACK_BOT_TOKEN }}
```

## Best Practices

1. **Channel Naming**: Use consistent, descriptive channel names
2. **Message Format**: Keep messages concise and scannable
3. **Threading**: Use threads for related updates
4. **Mentions**: Use `@channel` sparingly, prefer `@here` or specific users
5. **Emoji**: Use standard emoji for status indicators
6. **Rate Limits**: Be mindful of Slack API rate limits
7. **Error Handling**: Always handle failures gracefully

## Troubleshooting

### Token Issues

If Slack notifications fail:

1. Verify token is correct and starts with `xoxb-`
2. Check bot has required permissions
3. Ensure app is installed in workspace
4. Test token with curl:

```bash
curl -H "Authorization: Bearer $SLACK_BOT_TOKEN" \
  https://slack.com/api/auth.test
```

### Channel Not Found

1. Get correct channel ID (not name)
2. Ensure bot is member of private channels
3. Use `chat:write.public` scope for public channels

### Webhook Not Working

1. Verify webhook URL is complete
2. Check webhook hasn't been revoked
3. Test with curl:

```bash
curl -X POST -H 'Content-type: application/json' \
  --data '{"text":"Test message"}' \
  $SLACK_WEBHOOK_URL
```

## Resources

- [Slack API Documentation](https://api.slack.com/)
- [Block Kit Builder](https://app.slack.com/block-kit-builder) - Design interactive messages
- [Slack GitHub Action](https://github.com/marketplace/actions/slack-send)
- [Node Slack SDK](https://slack.dev/node-slack-sdk/)

## Next Steps

- [Review Enterprise Deployment Guide](../enterprise-deployment.md)
- [Configure GitHub Enterprise Integration](./github-enterprise.md)
- [Set up Atlassian Integration](./atlassian-integration.md)
