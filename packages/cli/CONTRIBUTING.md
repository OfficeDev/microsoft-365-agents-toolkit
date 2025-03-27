# Contributing to Microsoft 365 Agents Toolkit CLI 

Welcome, and thank you for your interest in contributing to Microsoft 365 Agents Toolkit CLI!

Please review this document for setting up your development environment, debugging and run Microsoft 365 Agents Toolkit CLI. If you have any questions, please raise your issue on github.

## Prerequisites
---

Verify you have the right prerequisites for building Teams apps:

### M365 account

The Microsoft 365 Agents Toolkit CLI requires a Microsoft 365 organizationl account where Teams is running and has been registered.

### Azure account
The Microsoft 365 Agents Toolkit CLI may require an Azure account and subscription to deploy the Azure resources for your project.

**_NOTE:_** Don't have a M365 to experience building Teams app? Sign up for [M365 Developer Program](https://docs.microsoft.com/en-us/microsoftteams/platform/concepts/build-and-test/prepare-your-o365-tenant), which allows you to have a testing tenant with preconfigured permissions.

## Setup Development Environment
---

1. Install Node v18 or later
2. Install PNPM v8 or later

## Build the project
---

### Build the whole monorepo
1. `git clone https://github.com/OfficeDev/TeamsFx.git`
2. `cd TeamsFx`
3. `npm run setup`

This will run "pnpm install && npm run build" to link packages in monorepo locally. 


## Debug the project
---

### Run your local project
1. `cd TeamsFx`
2. `npm run setup`
3. `npm link`

### Debug inside VSCode
1. `cd TeamsFx/packages/cli`
2. `code .`
3. In the debug Treeview choose debugging profile
4. Hit 'F5' or click start debugging button

### Install the published package
1. Run: `npm install -g @microsoft/m365agentstoolkit-cli` (Pls check the version is the latest version)
2. Now the package is installed in your global npm folder. You can type 'm365agents -h' to see how to use the cli tool.

## Test the proejct
---

### Run Unit test

1. `cd TeamsFx/packages/cli`
2. `npm run test:unit`

### Run E2E test

1. `cd TeamsFx/packages/cli`
2. `npm run test:e2e`

**_NOTE:_** When run E2E tests it may pop up windows to ask you to login your Azure account or M365 account, please login your accout to continue the test.

## Coding Style
---

The project setup ESLINT and prettier for coding style and formating, please follow the commands below

### Lint project
`npm run lint`

### Fix lint error
`npm run lint:fix`

### Check code format before commit
`npm run check-format`

### Format the code
`npm run format`

### Add dependency for CLI
`pnpm install XXX`

### Delete dependency for CLI
`pnpm remove XXX`

## Opening PR and PR review
---


# Thank You!

Your contributions to open source, large or small, make great projects like this possible. Thank you for taking the time to contribute.
