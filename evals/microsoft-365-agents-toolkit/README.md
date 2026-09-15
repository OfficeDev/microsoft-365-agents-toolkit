# Microsoft 365 Agents Toolkit Evaluation

This directory contains the evaluation framework for the Microsoft 365 Agents Toolkit skill.

## Overview

The evaluation setup is designed to test the skill's ability to guide users through various workflows:
- **Scaffolding**: Create new projects (bot, tab, message extension, declarative agent, custom engine agent)
- **Debugging**: Run projects locally in Agents Playground
- **Deployment**: Deploy projects to Azure
- **Negative triggers**: Verify the skill doesn't trigger for unrelated requests

## Structure

```
evals/microsoft-365-agents-toolkit/
├── eval.yaml                 # Main evaluation configuration
├── setup.sh                  # Setup script for CI environment
├── scaffold-tasks/           # Project creation test tasks
├── debug-tasks/              # Local debugging test tasks
├── deploy-tasks/             # Azure deployment test tasks
├── negative-tasks/           # Negative trigger test tasks
└── fixtures/                 # Test project files
    └── simple-bot/           # Sample bot project used in tests
```

## Running Evaluations

### Locally

To run the evaluation locally, you need:
1. Required environment variables (see setup.sh)
2. Installed tooling: `waza` CLI
3. For Declarative Agent execution tasks, WIQD installed with its EULA accepted interactively
4. Proper authentication set up for tasks that access Microsoft 365 or Azure

Confirm WIQD is installed and its EULA has been accepted before running Declarative Agent execution tasks:

```bash
wiqd --version
```

For tasks that access Microsoft 365, also verify WIQD authentication:

```bash
wiqd auth status
wiqd doctor
```

If WIQD is missing or its EULA has not been accepted, `setup.sh` emits a warning. Non-DA tasks can continue, but DA execution tasks will not pass. Read-only DA schema tasks do not require WIQD.

```bash
cd evals/microsoft-365-agents-toolkit
bash setup.sh
waza run eval.yaml
```

### Via GitHub Actions

The evaluation runs automatically on:
1. **Manual trigger**: Use the GitHub Actions UI with "Run workflow"
2. **Skill changes**: Push to `packages/vscode-extension/skills/microsoft-365-agents-toolkit/`
3. **Evaluation changes**: Push to `evals/microsoft-365-agents-toolkit/`

Required secrets in the repository:
- `M365_ACCOUNT_NAME` - Microsoft 365 test account
- `M365_ACCOUNT_PASSWORD` - Microsoft 365 account password
- `M365_TENANT_ID` - Microsoft 365 tenant ID
- `AZURE_ACCOUNT_NAME` - Azure test account
- `AZURE_ACCOUNT_OBJECT_ID` - Azure account object ID
- `AZURE_ACCOUNT_PASSWORD` - Azure account password
- `AZURE_TENANT_ID` - Azure tenant ID
- `AZURE_SUBSCRIPTION_ID` - Azure subscription ID

The schema-validation job runs on a fresh GitHub-hosted runner and does not require WIQD. Behavioral DA execution requires an environment where WIQD has already been initialized. WIQD does not currently provide a supported non-interactive EULA acceptance flow, so a fresh GitHub-hosted runner cannot bootstrap this prerequisite. Do not bypass or automate EULA acceptance; use a preconfigured runner when that becomes available and track the CI bootstrap dependency separately.

## Skill Location

The skill being evaluated is located at:
- [`packages/vscode-extension/skills/microsoft-365-agents-toolkit/`](../../packages/vscode-extension/skills/microsoft-365-agents-toolkit/)

## Metrics

The evaluation measures two key metrics:

1. **Task Completion (70% weight)** - Does the skill successfully guide users through the requested workflow?
   - Trigger and anti-trigger checks pass
   - Threshold: 80%

2. **Efficiency (30% weight)** - Does the skill stay within reasonable resource limits?
   - Tool call limits (max 15-40 depending on task)
   - Duration limits (max 150-350 seconds depending on task)
   - No forbidden operations
   - Threshold: 70%

## Adding New Tests

To add a new test task:

1. Create a YAML file in the appropriate directory (`scaffold-tasks/`, `debug-tasks/`, etc.)
2. Define the task structure with:
   - `id`: Unique identifier
   - `name`: Human-readable name
   - `inputs`: Test input (prompt and optional files)
   - `expected`: Expected outcome (should_trigger: true/false)
   - `graders`: Evaluation criteria (file checks, behavior, prompt grading)
3. Reference the task file in `eval.yaml` under `tasks`

See existing task files for examples.
