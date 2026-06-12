# Contributing

Welcome and thank you for your interest in contributing to **fx-core**! Before contributing to this project, please review this document for policies and procedures which will ease the contribution and review process for everyone. If you have questions, please raise your issue on github.

## Setup develop environment

Follow the official documents to install the required softwares:

1. [Git](https://git-scm.com/)
2. [Node.js and NPM](https://nodejs.org/), **x64**, Node version >= 22
3. [PNPM](https://pnpm.io/), PNPM version >=8
4. [Visual Studio Code](https://code.visualstudio.com/)


## Built the project

1. Clone this repo locally. (`git clone https://github.com/OfficeDev/TeamsFx.git`)
2. Open a terminal and move into your local copy. (`cd TeamsFx`)
3. Because the monorepo is managed by PNPM, you need to setup at the first time in root folder. (`npm run setup` or `pnpm install && npm run build`) All dependencies will be installed and linked locally.

## Add or remove dependency in FxCore
run `pnpm install XXX` to add dependency, run `pnpm remove XXX` to remove dependecy.

## Test the project

Add tests under tests/ folder. The filename should end with .test.ts.

Because other packages depends on `fx-core`, the change may break functionalities of other packages.
Please also run `pnpm -r run test:unit` in the root folder of TeamsFx project, it will run all unit tests.

## Style Guidelines

After running `npm run setup` in the root folder of the monorepo, a [Precommit Git Hook](https://git-scm.com/book/en/v2/Customizing-Git-Git-Hooks) will be added by [Husky](https://github.com/typicode/husky), which will run

1. [ESLint](https://github.com/eslint/eslint): Please fix ESLint errors & warning before committing your code.
2. [Prettier](https://github.com/prettier/prettier): Your code will be formatted automatically.

## Pull Request Process

1. Checkout your local branch from the latest `dev` branch and make your changes to your local branch.
2. Before creating a pull request, make sure:
   - Use eslint plugin to check whether there is any error or warning that breaks the rule. (`npm run lint`)
   - Make sure modified functions are covered by unit tests. (`npm run test`)
   - Better to add comment for public class/method.
3. Push your local branch and create the pull request.
4. Make sure all the checks in pull request are passed.
5. At least one approve from each code owner is required.
