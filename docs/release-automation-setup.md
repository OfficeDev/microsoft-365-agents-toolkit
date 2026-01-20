# Release Automation Setup Guide

This guide explains how to configure the release automation system.

## Prerequisites

### 1. GitHub Environment Setup

Create a GitHub environment named `release-approval`:

1. Go to repository **Settings** → **Environments**
2. Click **New environment**
3. Name it: `release-approval`
4. Add **Required reviewers** (team members who can approve releases)
5. Save

### 2. Email Notification Configuration

The workflow sends email notifications after triggering CD pipelines. Configure the following:

#### GitHub Variables

Set in **Settings** → **Secrets and variables** → **Actions** → **Variables**:

```
RELEASE_NOTIFICATION_EMAIL = "your-team@example.com,release-manager@example.com"
```

#### GitHub Secrets

The email action requires these secrets (should already exist if cd.yml is working):

```
MAIL_CLIENT_ID = "your-azure-ad-app-client-id"
MAIL_CLIENT_SECRET = "your-azure-ad-app-client-secret"  
MAIL_TENANT_ID = "your-azure-ad-tenant-id"
```

### 3. GitHub Token

The workflow automatically uses `secrets.GITHUB_TOKEN` for:
- Creating branches
- Triggering workflows
- Calling GitHub Models API for AI-powered parsing

No additional configuration needed - this is provided by GitHub Actions.

## Features

### 📄 Schedule-Driven Parsing

The system reads release parameters directly from the release schedule markdown table.
No AI-based inference is used.

### 🔒 Safety Features

- **Dry run mode** by default - review before executing
- **Manual approval** via `release-approval` environment before triggering CD
- **Detailed logs** showing all parameters before execution
- **Automatic skip** of already-released or canceled releases

### 📧 Email Notifications

After triggering CD pipeline, the team receives an email with:
- Product and version information
- Branch name
- CD parameters (preid, series, pkgs, vsrelease)
- Links to workflow runs
- Instructions to monitor progress

### ✅ Branch Management

- **Creates release branches** from dev if they don't exist
- **Ensures CD runs on release branch** (not dev)
- **Follows naming conventions**:
  - VS Code: `release/6.5`
  - Visual Studio: `release/VS180I4`

## Usage

### Manual Execution

1. Go to **Actions** → **Release Automation**
2. Click **Run workflow**
3. Configure options:
   - `release_index`: Which release to process (see step 4)
   - `dry_run`: Set to `false` to execute (default: `true`)
   - `create_branch`: Whether to create branch if missing (default: `true`)
4. First run in dry mode to see available releases and their indexes
5. Run again with specific index and `dry_run=false`
6. If a sync PR (`dev` → release branch) is created, merge it first
7. Approve in the `release-approval` environment gate (CD triggers after approval)
8. Monitor execution

### Scheduled Checks

The workflow automatically runs **every Monday at 8 AM UTC** to:
- Parse the release schedule
- Identify pending releases
- Create a GitHub issue with review information
- Team reviews and manually triggers for specific releases

## Parameter Source

Release parameters are taken from the schedule table columns.

Required columns:
- `Branch`
- `preid`
- `series`
- `pkgs`
- `vsrelease`

## Troubleshooting

### Parsing Issues

If a release entry is missing required columns, it will be skipped and the job log will include an error.

### Email Not Sending

1. Verify email variables are set correctly
2. Check that Azure AD app credentials are valid
3. Review email action logs for errors
4. Test with cd.yml email notifications first

### Branch Not Created

1. Ensure `create_branch` is set to `true`
2. Check that the workflow has write permissions to repository
3. Verify the GitHub App token has necessary permissions

## Customization

To change parameters, update the corresponding values in the schedule markdown file.

### Modify Email Template

Edit the `BODY` environment variable in the workflow file to customize the email content.

### Add More Recipients

Update the `RELEASE_NOTIFICATION_EMAIL` variable with comma-separated email addresses.

## File Structure

```
.github/
├── workflows/
│   ├── cd.yml                           # Main CD pipeline
│   └── release-automation.yml           # New automation workflow
├── scripts/
│   ├── parse_release_schedule.py        # Original rule-based parser (backup)
│   └── parse_release_schedule_ai.py     # New AI-powered parser
└── actions/
    └── send-email-report/              # Email notification action

docs/
└── release-schedule.md                  # Release schedule source
```
