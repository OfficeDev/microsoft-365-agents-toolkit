# Environment Configuration Reference

## Contents
- Environment YAML Files (which YAML maps to which env)
- .localConfigs vs env/.env.local (critical runtime config flow)
- Missing Environment Variables at Runtime
- Azure OpenAI Configuration
- YAML File Structure (m365agents.yml)
- Cloud Deployment Workflow
- Key Environment Variables

## Environment YAML Files (Lifecycle Configuration)

`atk provision` and `atk deploy` execute actions defined in environment-specific YAML files:

| YAML File | Used When | Purpose |
|-----------|-----------|---------|
| `m365agents.yml` | `--env dev` (default) | Production/dev deployment to Azure |
| `m365agents.local.yml` | `--env local` | Local development setup |
| `m365agents.playground.yml` | Agents Playground testing | Test with Agents Playground |
| `teamsapp.yml` | Legacy projects | Older Teams Toolkit format |

```bash
# Runs actions in m365agents.yml (Azure + M365 resources)
atk provision --env dev --resource-group <rg> --region <region> -i false
atk deploy --env dev -i false

# Runs actions in m365agents.local.yml
atk provision --env local -i false
```

## .localConfigs vs env/.env.local

For `local` and `playground` environments, the YAML uses `file/createOrUpdateEnvironmentFile` to write environment variables to `.localConfigs` (not `env/.env.local`).

```yaml
# Example from m365agents.local.yml
- uses: file/createOrUpdateEnvironmentFile
  with:
    target: ./.localConfigs
    envs:
      BOT_ID: ${{BOT_ID}}
      BOT_PASSWORD: ${{SECRET_BOT_PASSWORD}}
```

**Critical:** The backend service reads from `.localConfigs`, NOT from `env/.env.local`.

**Configuration flow:**
- `env/.env.local` → source of truth (edited manually or by `atk provision`)
- `m365agents.local.yml` → defines how to transform env vars
- `.localConfigs` → generated file your app reads at runtime (created by `atk deploy --env local`)

## Missing Environment Variables at Runtime

If your app is missing environment variables:
1. Check `.localConfigs` exists and has the required values
2. Ensure values are set in `env/.env.local` (or `env/.env.local.user` for secrets)
3. Ensure `m365agents.local.yml` maps those values to `.localConfigs`
4. Run `atk deploy --env local -i false` to regenerate `.localConfigs`

## Azure OpenAI Configuration

For custom engine agents using Azure OpenAI, add these env vars to the YAML's `file/createOrUpdateEnvironmentFile` action:

```yaml
# Add to m365agents.local.yml or m365agents.playground.yml
- uses: file/createOrUpdateEnvironmentFile
  with:
    target: ./.localConfigs
    envs:
      AZURE_OPENAI_API_KEY: ${{SECRET_AZURE_OPENAI_API_KEY}}
      AZURE_OPENAI_ENDPOINT: ${{AZURE_OPENAI_ENDPOINT}}
      AZURE_OPENAI_DEPLOYMENT_NAME: ${{AZURE_OPENAI_DEPLOYMENT_NAME}}
```

Then set values in `env/.env.local`:
```
SECRET_AZURE_OPENAI_API_KEY=your-api-key
AZURE_OPENAI_ENDPOINT=https://your-resource.openai.azure.com/
AZURE_OPENAI_DEPLOYMENT_NAME=gpt-4o
```

After updating the YAML, run `atk deploy --env local -i false` to write values to `.localConfigs`.

## YAML File Structure (m365agents.yml)

m365agents.yml handles both M365 resources AND Azure resources. For Azure provisioning/deployment (`arm/deploy`, `azureAppService/zipDeploy`, `azureFunctions/zipDeploy`), use `--resource-group` and `--region` with `atk provision`.

**Prerequisites for Azure resources:**
1. Set `AZURE_SUBSCRIPTION_ID` in `env/.env.dev`
2. Run `atk auth login azure`

```yaml
version: 1.0.0

provision:
  - uses: teamsApp/create        # Register app in Teams Developer Portal
  - uses: botAadApp/create       # Create bot Azure AD app
  - uses: arm/deploy             # Deploy Azure resources (requires --resource-group and --region)
  - uses: teamsApp/zipAppPackage # Package the app
  - uses: teamsApp/update        # Update app registration

deploy:
  - uses: cli/runNpmCommand      # Build the project
  - uses: azureAppService/zipDeploy  # Deploy to Azure App Service
```

## Cloud Deployment Workflow

```bash
# Step 0: Check required environment variables before provisioning
# Look at m365agents.yml for ${{VAR_NAME}} references
# Copy relevant values from env/.env.local or env/.env.playground to env/.env.dev

# Step 1: Configure Azure subscription in env/.env.dev
# AZURE_SUBSCRIPTION_ID=your-subscription-id

# Step 2: Create resource group if needed
# az group create --name <rg> --location <region>
# IMPORTANT: Verify az account matches atk account: `az account show` vs `atk auth list`

# Step 3: Provision Azure + M365 resources
atk provision --env dev --resource-group <rg> --region <region> -i false

# Step 4: Deploy code to Azure
atk deploy --env dev -i false

# Step 5: Open app in Teams/M365
# https://teams.microsoft.com/l/app/${{TEAMS_APP_ID}}?installAppPackage=true&webjoin=true&appTenantId=${{TENANT_ID}}&login_hint=${{USER_EMAIL}}
```

**Environment variables are stored in:**
- `env/.env.dev` - Non-secret configuration
- `env/.env.dev.user` - Secrets (prefixed with `SECRET_`)

**Customizing environments:**
- Edit the YAML to add/remove/modify lifecycle actions
- Each action (`uses:`) corresponds to a built-in ATK task
- Environment variables from `env/.env.<envname>` are available during execution

## Key Environment Variables

| Variable | Description |
|----------|-------------|
| `TEAMS_APP_ID` | Teams app ID (for Teams bot/tab apps) |
| `M365_APP_ID` | M365 app ID (for Declarative Agents in Copilot) - acquired after `teamsApp/extendToM365` |
| `TENANT_ID` | Azure AD tenant ID |
| `BOT_ID` | Bot Azure AD app ID |
| `SECRET_BOT_PASSWORD` | Bot Azure AD app secret |
| `AZURE_SUBSCRIPTION_ID` | Azure subscription (for cloud deployment) |

**Important:** Declarative agents use `M365_APP_ID`, not `TEAMS_APP_ID`. If `M365_APP_ID` is missing, the agent won't load in M365 Copilot.
