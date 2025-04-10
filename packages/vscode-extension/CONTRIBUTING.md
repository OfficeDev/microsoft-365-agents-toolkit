# Contributing

Welcome and thank you for your interest in contributing to **VSCode Extension**! Before contributing to this project, please review this document for policies and procedures which will ease the contribution and review process for everyone. If you have questions, please raise your issue on github.

## Setup develop environment

### Prerequisites

Verify you have the right prerequisites for building Teams apps and install some recommended development tools. [Read more details](https://docs.microsoft.com/en-us/microsoftteams/platform/build-your-first-app/build-first-app-overview).



<table>
    <tr>
        <td><h3>Node.js and PNPM</h3>As a fundamental runtime context for Teams app, Node.js v18 or higher, PNPM v8 or higher.</td>
    </tr>
    <tr>
        <td><h3>Microsoft 365</h3>The Teams Toolkit requires a Microsoft 365 organizational account where Teams is running and has been registered.</td>
    </tr>
    <tr>
        <td><h3>Azure</h3> The Teams Toolkit may require an Azure account and subscription to deploy the Azure resources for your project.</td>
    </tr>
</table>

> Don’t have a Microsoft 365 account to experience building Teams app? Sign up for [Microsoft Developer Program](https://developer.microsoft.com/en-us/microsoft-365/dev-program), which allows you to have a testing tenant with preconfigured permissions.

### Built the project

1. Clone this repo locally. (`git clone https://github.com/OfficeDev/TeamsFx.git`)
1. Open a terminal and move into your local copy. (`cd TeamsFx`)
1. Because the monorepo is managed by PNPM, you need to setup the project at the first time. (`npm run setup` or `pnpm install && npm run build` at root folder) All dependencies will be installed.

### Debug the project

1. Open project in VS Code (`cd packages/vscode-extension && code .`)
1. Press `F5` in VS Code.

### Add or remove dependecy in vscode-extension
run `pnpm install XXX` to add dependency, run `pnpm remove XXX` to remove dependency.

## Test the project

Mannully test UI in VS Code extenion for now.

## Style Guidelines

The project already enabled StyleCop. Please fix the style warnings before commit.

## Pull Request Process

1. Check out a new branch from `dev` branch.
1. Make sure all the checks in pull request are passed.
1. At least one approver from [CODEOWNER](../../.github/CODEOWNERS) is required.
