# Getting started

## Prerequisites

- Node.js 18.x or 20.x
- pnpm 9.x (`npm i -g pnpm`)
- Git
- For .NET SDK / function-extension work: .NET 8 SDK
- For Python templates: Python 3.10+

## Clone and bootstrap

```bash
git clone https://github.com/OfficeDev/teams-toolkit.git
cd teams-toolkit
pnpm install                 # installs all workspaces
npm run setup                # full build (slow first time)
```

For a faster first build, target only what you need:

```bash
npm run setup:cli            # CLI v3 + dependencies
npm run setup:vsc            # VS Code extension + dependencies
pnpm --filter ./packages/core-next run build
pnpm --filter ./packages/cli-next  run build
```

## CLA

First contribution requires signing the [Microsoft CLA](https://cla.opensource.microsoft.com). The CLA bot will instruct you on the PR.

## Code of Conduct

This project follows the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/).

## Where to learn the codebase

1. Skim this site's [00-overview](../00-overview/README.md).
2. Read the [coding standards](../05-engineering/coding-standards.md).
3. Pick a contributor playbook below.

## Common first contributions

- Add a [template](adding-a-template.md).
- Add a [driver](adding-a-driver.md).
- Add a [CLI command](adding-a-cli-command.md).
- Improve a [troubleshooting page](../08-troubleshooting/README.md).
- Translate strings (see [`Localize/loc/`](../../Localize/loc/)).
