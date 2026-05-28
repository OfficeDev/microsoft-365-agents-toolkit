// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { NextStep } from "../../../chat/commands/nextstep/types";
import { CHAT_EXECUTE_COMMAND_ID } from "../../../chat/consts";
import { CommandKey } from "../../../constants";
import {
  canOfficeAddInPreviewInLocalEnv,
  isDebugSucceededAfterSourceCodeChanged,
  isDependenciesInstalled,
  isNodeInstalled,
  isProjectOpened,
} from "./condition";
import { OfficeWholeStatus } from "./types";

// TODO: align the description with PM
export const officeSteps: () => NextStep[] = () => [
  {
    title: "Create a New Project",
    description:
      "To get started, you can create a new Office Add-in project by using `/create` to build your Office add-in as per your description.",
    docLink: "",
    commands: [],
    followUps: [
      {
        label: "@office /create an Excel hello world add-in",
        command: "create",
        prompt: "an Excel hello world add-in",
      },
      {
        label: "@office /create a Word add-in that inserts comments",
        command: "create",
        prompt: "a Word add-in that inserts comments",
      },
    ],
    condition: (status: OfficeWholeStatus) => !isProjectOpened(status),
    priority: 0,
  },
  {
    title: "Check Prerequisites",
    description: `To get ready for add-in development, you need to have [Node.js v16/v18](https://nodejs.org/) and [npm](https://www.npmjs.com/get-npm) installed. You can check your environment by clicking the button below.`,
    docLink:
      "https://learn.microsoft.com/en-us/office/dev/add-ins/concepts/requirements-for-running-office-add-ins",
    commands: [
      {
        title: "Check prerequisites",
        command: CHAT_EXECUTE_COMMAND_ID,
        arguments: [CommandKey.ValidateGetStartedPrerequisites],
      },
    ],
    followUps: [],
    condition: (status: OfficeWholeStatus) => isProjectOpened(status) && !isNodeInstalled(status),
    priority: 1,
  },
  {
    title: "Install Dependencies",
    description: `Install the dependencies for your Office Add-ins project. It runs ''npm install'' command to install all the dependencies in the terminal.`,
    docLink: "",
    commands: [
      {
        title: "Install Dependencies",
        command: CHAT_EXECUTE_COMMAND_ID,
        arguments: [CommandKey.installDependency],
      },
    ],
    followUps: [],
    condition: (status: OfficeWholeStatus) =>
      isProjectOpened(status) && isNodeInstalled(status) && !isDependenciesInstalled(status),
    priority: 1,
  },
  {
    title: "Preview in Local Environment",
    description: `To run and debug the add-in, you can preview the add-in in Office apps to understand how it works. Start debugging by clicking the button below or pressing \`F5\`.\n\nIf you meet problems, please check the \`README.md\` file for detailed guidance. `,
    docLink: "https://learn.microsoft.com/en-us/office/dev/add-ins/testing/debug-add-ins-overview",
    commands: [
      {
        title: "Preview in Local Environment",
        command: CHAT_EXECUTE_COMMAND_ID,
        arguments: [CommandKey.LocalDebug],
      },
    ],
    followUps: [],
    condition: (status: OfficeWholeStatus) =>
      isProjectOpened(status) &&
      isNodeInstalled(status) &&
      isDependenciesInstalled(status) &&
      canOfficeAddInPreviewInLocalEnv(status) &&
      !isDebugSucceededAfterSourceCodeChanged(status),
    priority: 1,
  },
  {
    title: "Code Generation",
    description:
      "To customize the add-in project, you can generate code for Office add-ins by using `/generatecode` to describe the feature you would like to build.",
    docLink: "",
    commands: [],
    followUps: [
      {
        label: "@office /generatecode create a chart based on the selected range in Excel",
        command: "generatecode",
        prompt: "create a chart based on the selected range in Excel",
      },
      {
        label: "@office /generatecode insert a content control in a Word document",
        command: "generatecode",
        prompt: "insert a content control in a Word document",
      },
    ],
    condition: (status: OfficeWholeStatus) =>
      isProjectOpened(status) &&
      isNodeInstalled(status) &&
      isDependenciesInstalled(status) &&
      isDebugSucceededAfterSourceCodeChanged(status),
    priority: 1,
  },
  {
    title: "Deploy or Publish",
    description: `To distribute your add-in to a wider audience, you can [deploy](https://learn.microsoft.com/en-us/office/dev/add-ins/publish/publish#deployment-options-by-office-application-and-add-in-type) or [publish](https://learn.microsoft.com/en-us/office/dev/add-ins/publish/publish-office-add-ins-to-appsource) it.`,
    docLink: "",
    commands: [],
    followUps: [],
    condition: (status: OfficeWholeStatus) =>
      isProjectOpened(status) &&
      isNodeInstalled(status) &&
      isDependenciesInstalled(status) &&
      isDebugSucceededAfterSourceCodeChanged(status),
    priority: 2,
  },
];
