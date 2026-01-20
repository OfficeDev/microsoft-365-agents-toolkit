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

### 🤖 AI-Powered Parsing

The system uses **GitHub Models API** (GPT-4o) to intelligently parse release schedules:

- **Robust parsing**: Handles variations in version formats, dates, and product names
- **Automatic fallback**: Falls back to rule-based parsing if AI is unavailable
- **Smart parameter generation**: Determines branch names, preid, series, etc.

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

## Parameter Rules

The AI model follows these rules (with fallback to hard-coded logic):

### Branch Names

- **VS Code**: `release/X.Y` (major.minor)
  - Example: `release/6.5` for versions 6.5.0, 6.5.20251203
  
- **Visual Studio**: `release/VSXXYTZ`
  - XX = major version
  - Y = minor version  
  - T = type (P=Preview, I=Insider)
  - Z = number
  - Example: `release/VS180I4` for "18.0 Insider 4"

### Preid

- **VS Code**:
  - `preview` if minor version is odd (6.3, 6.5) OR release type is "Prerelease"
  - `rc` if minor version is even (6.0, 6.2, 6.4) AND release type is "Stable"
  
- **Visual Studio**: Always `preview`

### Series

- **VS Code**: `CYYMMDD` format from cut date
  - Example: `CY260106` for January 6, 2026
  
- **Visual Studio**: Same as branch name without `release/` prefix
  - Example: `VS180I4`

### Packages & VS Release

- **VS Code**: `pkgs=""`, `vsrelease="false"`
- **Visual Studio**: `pkgs="server"`, `vsrelease="true"`

## Troubleshooting

### AI Parsing Not Working

If GitHub Models API is unavailable:
- System automatically falls back to rule-based parsing
- Check GitHub Actions logs for warnings
- Ensure `GITHUB_TOKEN` has proper permissions

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

### Change AI Model

Edit `.github/scripts/parse_release_schedule_ai.py`:

```python
def call_github_model(prompt: str, model: str = "gpt-4o-mini") -> str:
    # Change model parameter default value
```

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
