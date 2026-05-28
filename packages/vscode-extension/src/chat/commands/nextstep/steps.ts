// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { CommandKey } from "../../../constants";
import { CHAT_EXECUTE_COMMAND_ID, CHAT_OPENURL_COMMAND_ID } from "../../consts";
import {
  canPreviewInTestTool,
  isAzureAccountLogin,
  isDebugSucceededAfterSourceCodeChanged,
  isDeployedAfterSourceCodeChanged,
  isDidNoActionAfterScaffolded,
  isFirstInstalled,
  isHaveReadMe,
  isM365AccountLogin,
  isProjectOpened,
  isProvisionedSucceededAfterInfraCodeChanged,
  isPublishedSucceededBefore,
} from "./condition";
import { NextStep, WholeStatus } from "./types";

export const allSteps: () => NextStep[] = () => [
  {
    title: "Getting started with Teams Toolkit",
    description: `Teams Toolkit makes it simple to get started with app development for Microsoft Teams using Visual Studio Code. You can start with a project template for a common custom app built for your org (LOB app) scenarios or from a sample. You can save setup time with automated app registration and configuration. You can run and debug your app in Teams directly from familiar tools. You can smart defaults for hosting in Azure using infrastructure-as-code and Bicep. You can create unique configurations like dev, test, and prod using the environment features.`,
    docLink: "https://aka.ms/vsc-ttk-v5-install",
    commands: [
      {
        title: "Open Get-Started Page",
        command: CHAT_EXECUTE_COMMAND_ID,
        arguments: [CommandKey.OpenWelcome],
      },
      {
        title: "Open Document",
        command: CHAT_EXECUTE_COMMAND_ID,
        arguments: [CommandKey.OpenDocument],
      },
    ],
    followUps: [
      {
        label: "@teams /create",
        command: "create",
        prompt: "",
      },
    ],
    condition: (status: WholeStatus) => isFirstInstalled(status),
    priority: 0,
  },
  {
    title: "Create a new project or open an existing project",
    description:
      "You have the option to create a new Teams Toolkit project, or open one that already exists. For new projects, you can start with @teams /create to use the built-in Teams app templates or click the button below to use the official Teams app samples in Teams Toolkit. Also, Teams Toolkit v5 allows you to use Outlook Add-in templates to build your own Outlook Add-ins.",
    docLink: "https://aka.ms/teamsfx-create-project",
    commands: [
      {
        title: "Open Sample Gallery",
        command: CHAT_EXECUTE_COMMAND_ID,
        arguments: [CommandKey.OpenSamples],
      },
      {
        title: "Create Teams App: Find Templates/Samples",
        command: CHAT_EXECUTE_COMMAND_ID,
        arguments: [
          "workbench.action.chat.open",
          { query: "@teams /create ", isPartialQuery: true },
        ],
      },
    ],
    followUps: [],
    condition: (status: WholeStatus) => !isProjectOpened(status),
    priority: 0,
  },
  {
    title: "Get more info about the project with README",
    description: (status: WholeStatus) => {
      // readme must exist because the condition has checked it
      const readme = status.projectOpened!.readmeContent!;
      let description = "";
      let findFirstSharp = false;
      for (const line of readme.split("\n")) {
        if (findFirstSharp && line.trim().startsWith("#")) {
          break;
        }
        if (line.trim().startsWith("#")) {
          findFirstSharp = true;
        }
        if (!findFirstSharp) {
          continue;
        }
        description += line.trim() + " ";
      }
      return description;
    },
    commands: [
      {
        title: "View Full README",
        command: CHAT_EXECUTE_COMMAND_ID,
        arguments: [CommandKey.OpenReadMe],
      },
    ],
    followUps: [],
    condition: (status: WholeStatus) =>
      isProjectOpened(status) && isDidNoActionAfterScaffolded(status) && isHaveReadMe(status),
    priority: 1,
  },
  {
    title: "Preview in Test Tool",
    description: `Teams App Test Tool (Test Tool) makes debugging bot-based apps effortless. You can chat with your bot and see its messages and Adaptive Cards as they appear in Teams. You don't need a Microsoft 365 developer account, tunneling, or Teams app and bot registration to use Test Tool. When previewing with Test Tool, it will check all required prerequisites and guide you to fix them in output.`,
    docLink: "https://aka.ms/vsc-ttk-teams-app-test-tool",
    commands: [
      {
        title: "Preview in Test Tool",
        command: CHAT_EXECUTE_COMMAND_ID,
        arguments: [CommandKey.DebugInTestToolFromMessage],
      },
    ],
    followUps: [],
    condition: (status: WholeStatus) =>
      isProjectOpened(status) &&
      !isDidNoActionAfterScaffolded(status) &&
      !isDebugSucceededAfterSourceCodeChanged(status) &&
      canPreviewInTestTool(status),
    priority: 0,
  },
  {
    title: "Sign in to Microsoft 365 Account",
    description: `Preview in Teams requires a Microsoft 365 developer account. If you have a Visual Studio Enterprise or Professional subscription, both programs include a free Microsoft 365 developer subscription. It's active as long as your Visual Studio subscription is active.`,
    docLink: "https://aka.ms/vsc-ttk-m365-dev-program",
    commands: [
      {
        title: "Sign in to Microsoft 365 Account",
        command: CHAT_EXECUTE_COMMAND_ID,
        arguments: [CommandKey.SigninM365],
      },
    ],
    followUps: [],
    condition: (status: WholeStatus) =>
      isProjectOpened(status) &&
      !isDidNoActionAfterScaffolded(status) &&
      !isDebugSucceededAfterSourceCodeChanged(status) &&
      !isM365AccountLogin(status),
    priority: 1,
  },
  {
    title: "Join Microsoft 365 Developer Program",
    description: `If you don't have any Microsoft 365 tenant, you might qualify for a Microsoft 365 E5 developer subscription through the Microsoft 365 Developer Program; Alternatively, you can sign up for a 1-month free trial or purchase a Microsoft 365 plan.`,
    docLink: "https://aka.ms/vsc-ttk-create-m365-dev-account",
    commands: [
      {
        title: "Join Microsoft 365 Developer Program",
        command: CHAT_EXECUTE_COMMAND_ID,
        arguments: [
          CHAT_OPENURL_COMMAND_ID,
          "https://developer.microsoft.com/en-us/microsoft-365/dev-program",
        ],
      },
    ],
    followUps: [],
    condition: (status: WholeStatus) =>
      isProjectOpened(status) &&
      !isDidNoActionAfterScaffolded(status) &&
      !isDebugSucceededAfterSourceCodeChanged(status) &&
      !isM365AccountLogin(status),
    priority: 2,
  },
  {
    title: "Preview in Microsoft Teams",
    description: `Teams Toolkit helps you to debug and preview your Microsoft Teams app locally. During the debugging process, Teams Toolkit automatically starts app services, launches debuggers, and uploads Teams app. You can preview your Teams app in Teams web client locally after debugging. When previewing with Microsoft Teams, it will check all required prerequisites and guide you to fix them in output.`,
    docLink: "https://aka.ms/vsc-ttk-debug-local",
    commands: [
      {
        title: "Preview in Microsoft Teams",
        command: CHAT_EXECUTE_COMMAND_ID,
        arguments: [CommandKey.LocalDebug],
      },
    ],
    followUps: [],
    condition: (status: WholeStatus) =>
      isProjectOpened(status) &&
      !isDidNoActionAfterScaffolded(status) &&
      !isDebugSucceededAfterSourceCodeChanged(status) &&
      isM365AccountLogin(status),
    priority: 0,
  },
  {
    title: "How to Extend your Teams Application Capabilities",
    description: (status: WholeStatus) => {
      // readme must exist because the condition has checked it
      const readme = status.projectOpened!.readmeContent!;
      let description = "You can follow the README to extend the app, such as: ";
      for (const line of readme.split("\n")) {
        if (line.trim().startsWith("## Extend")) {
          description += line.trim().replace("##", "") + " ";
        }
      }
      return description;
    },
    commands: [
      {
        title: "Open README",
        command: CHAT_EXECUTE_COMMAND_ID,
        arguments: [CommandKey.OpenReadMe],
      },
    ],
    followUps: [],
    condition: (status: WholeStatus) =>
      isProjectOpened(status) &&
      !isDidNoActionAfterScaffolded(status) &&
      isDebugSucceededAfterSourceCodeChanged(status) &&
      isHaveReadMe(status),
    priority: 2,
  },
  {
    title: "Set up CI/CD Pipelines",
    description:
      "TeamsFx helps to automate your development workflow while building Teams application. The tools and templates to set up CI/CD pipelines are create workflow templates and customize CI/CD workflow with GitHub, Azure DevOps, Jenkins, and other platforms.",
    docLink: "https://aka.ms/teamsfx-add-cicd-new",
    commands: [],
    followUps: [], // TODO: point to S3
    condition: (status: WholeStatus) =>
      isProjectOpened(status) &&
      !isDidNoActionAfterScaffolded(status) &&
      isDebugSucceededAfterSourceCodeChanged(status),
    priority: 2,
  },
  {
    title: "Deploy Your App using Your Azure Account",
    description:
      "An Azure account allows you to host a Teams app or the back-end resources for your Teams app to Azure. You can do this using Teams Toolkit in Visual Studio Code. You must have an Azure subscription in the following scenarios: If you already have an existing app on a different cloud provider other than Azure, and you want to integrate the app on Teams platform. If you want to host the back-end resources for your app using another cloud provider, or on your own servers if they're available in the public domain.",
    docLink: "https://aka.ms/vsc-ttk-azure-account",
    commands: [
      {
        title: "Sign in to Azure Account",
        command: CHAT_EXECUTE_COMMAND_ID,
        arguments: [CommandKey.SigninAzure],
      },
    ],
    followUps: [],
    condition: (status: WholeStatus) =>
      isProjectOpened(status) &&
      !isDidNoActionAfterScaffolded(status) &&
      isDebugSucceededAfterSourceCodeChanged(status) &&
      !isProvisionedSucceededAfterInfraCodeChanged(status) &&
      !isAzureAccountLogin(status),
    priority: 1,
  },
  {
    title: "Provision Azure resources",
    description:
      "Teams Toolkit integrates with Azure and the Microsoft 365 cloud, which allows you to place your app in Azure with a single command. Teams Toolkit integrates with Azure Resource Manager (ARM) to set up Azure resources that your application needs, following a code-driven approach.",
    docLink: "https://aka.ms/teamsfx-provision",
    commands: [
      {
        title: "Provision Azure resources",
        command: CHAT_EXECUTE_COMMAND_ID,
        arguments: [CommandKey.Provision],
      },
    ],
    followUps: [],
    condition: (status: WholeStatus) =>
      isProjectOpened(status) &&
      !isDidNoActionAfterScaffolded(status) &&
      isDebugSucceededAfterSourceCodeChanged(status) &&
      !isProvisionedSucceededAfterInfraCodeChanged(status) &&
      isAzureAccountLogin(status),
    priority: 0,
  },
  {
    title: "Deploy to Azure",
    description: `Teams Toolkit helps to deploy or upload the front-end and back-end code in your app to your provisioned cloud resources in Azure. You can deploy to the following types of cloud resources: Azure App Services, Azure Functions, Azure Storage (as static website) and SharePoint`,
    docLink: "https://aka.ms/teamsfx-deploy",
    commands: [
      {
        title: "Deploy to Azure",
        command: CHAT_EXECUTE_COMMAND_ID,
        arguments: [CommandKey.Deploy],
      },
    ],
    followUps: [],
    condition: (status: WholeStatus) =>
      isProjectOpened(status) &&
      !isDidNoActionAfterScaffolded(status) &&
      isDebugSucceededAfterSourceCodeChanged(status) &&
      isProvisionedSucceededAfterInfraCodeChanged(status) &&
      !isDeployedAfterSourceCodeChanged(status),
    priority: 0,
  },
  {
    title: "Publish Your App",
    description:
      "After creating the app, you can distribute your app to different scopes, such as an individual, a team, or an organization. The distribution depends on multiple factors such as needs, business and technical requirements, and your goal for the app. Distribution to different scope may need different review processes. In general, the bigger the scope, the more review the app needs to go through for security and compliance concerns.",
    docLink: "https://aka.ms/teamsfx-publish",
    commands: [
      {
        title: "Publish Your App",
        command: CHAT_EXECUTE_COMMAND_ID,
        arguments: [CommandKey.Publish],
      },
    ],
    followUps: [],
    condition: (status: WholeStatus) =>
      isProjectOpened(status) &&
      !isDidNoActionAfterScaffolded(status) &&
      isDebugSucceededAfterSourceCodeChanged(status) &&
      isProvisionedSucceededAfterInfraCodeChanged(status) &&
      isDeployedAfterSourceCodeChanged(status) &&
      !isPublishedSucceededBefore(status),
    priority: 0,
  },
  {
    title: "Preview Remotely",
    description:
      "After provisioning and deploying the app to the remote, you can open the app in Teams client to see the real effect.",
    commands: [
      {
        title: "Preview Remotely",
        command: CHAT_EXECUTE_COMMAND_ID,
        arguments: [CommandKey.Preview],
      },
    ],
    followUps: [],
    condition: (status: WholeStatus) =>
      isProjectOpened(status) &&
      !isDidNoActionAfterScaffolded(status) &&
      isDebugSucceededAfterSourceCodeChanged(status) &&
      isProvisionedSucceededAfterInfraCodeChanged(status) &&
      isDeployedAfterSourceCodeChanged(status),
    priority: 1,
  },
];
