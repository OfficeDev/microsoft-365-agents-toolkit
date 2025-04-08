// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

"use strict";

import { FolderName } from "@microsoft/teamsfx-core";

export enum Browser {
  chrome = "chrome",
  edge = "edge",
  default = "default",
}

export class LaunchUrl {
  public static readonly teams: string =
    "https://teams.microsoft.com/l/app/${teamsAppId}?installAppPackage=true&webjoin=true&${account-hint}";
  public static readonly outlookTab: string =
    "https://outlook.office.com/host/${teamsAppInternalId}?${account-hint}";
  public static readonly outlookBot: string = "https://outlook.office.com/mail?${account-hint}";
  public static readonly officeTab: string =
    "https://www.office.com/m365apps/${teamsAppInternalId}?auth=2&${account-hint}";
}

export const teamsAppIdPlaceholder = "${teamsAppId}";
export const teamsAppInternalIdPlaceholder = "${teamsAppInternalId}";
export const accountHintPlaceholder = "${account-hint}";

export const serviceLogHintMessage = "The log of this task can be found in:";
export const openBrowserHintMessage =
  "WARN: Failed to open the browser, please copy the preview url and paste it into your browser.";
export const waitCtrlPlusC =
  "WARN: Closing browser will not terminate the preview process, please press Ctrl+C to terminate.";

export const frontendHostingPluginName = "fx-resource-frontend-hosting";
export const functionPluginName = "fx-resource-function";
export const botPluginName = "fx-resource-bot";
export const solutionPluginName = "solution";
export const appstudioPluginName = "fx-resource-appstudio";
export const spfxPluginName = "fx-resource-spfx";

export const teamsAppTenantIdConfigKey = "teamsAppTenantId";
export const remoteTeamsAppIdConfigKey = "teamsAppId";
export const botIdConfigKey = "botId";

export const frontendStartPattern = /Compiled|Failed/g;
export const backendStartPattern =
  /Worker process started and initialized|Host lock lease acquired by instance ID/g;
// From vscode $tsc-watch problem matcher: https://github.com/microsoft/vscode/blob/5a0ab56492d0c99f08028ca62ac3d59edb37f30f/extensions/typescript-language-features/package.json#L1085
export const tscWatchPattern =
  /^\s*(?:message TS6042:|\[?\D*\d{1,2}[:.]\d{1,2}[:.]\d{1,2}\D*(├\D*\d{1,2}\D+┤)?(?:\]| -)) (?:Compilation complete\.|Found \d+ errors?\.) Watching for file changes\./g;
// make a copy to prevent accidental change
export const backendWatchPattern = new RegExp(tscWatchPattern);
export const funcHostedBotWatchPattern = new RegExp(tscWatchPattern);
export const authStartPattern = /.*/g;
export const ngrokStartPattern = /started tunnel|failed to reconnect session/g;
export const botStartPattern = /listening|[nodemon] app crashed/g;
export const funcHostedBotStartPattern =
  /Worker process started and initialized|Host lock lease acquired by instance ID/g;
export const funcHostedBotAzuritePattern = /successfully listening/g;
export const gulpServePattern = /^.*Finished subtask 'reload'.*/g;
export const defaultRunningPattern =
  /started|successfully|finished|crashed|failed|listening|ready/i;

export const spfxInstallStartMessage = `executing 'npm install' under ${FolderName.SPFx} folder.`;
export const gulpCertTitle = "gulp trust-dev-cert";
export const gulpCertStartMessage = `executing 'gulp trust-dev-cert' under ${FolderName.SPFx} folder.`;
export const gulpServeTitle = "gulp serve";
export const gulpServeStartMessage = `executing 'gulp serve' under ${FolderName.SPFx} folder.`;
export const frontendInstallStartMessage = `executing 'npm install' under ${FolderName.Frontend} folder.`;
export const frontendStartStartMessage = `executing 'react-scripts start' under ${FolderName.Frontend} folder.`;
export const frontendStartStartMessageNext = `executing 'npm run dev:teamsfx' under ${FolderName.Frontend} folder.`;
export const authStartStartMessage = "starting auth service.";
export const backendInstallStartMessage = `executing 'npm install' under ${FolderName.Function} folder.`;
export const backendExtensionsInstallStartMessage =
  "installing Azure Functions binding extensions.";
export const backendStartStartMessage = `executing 'func start' under ${FolderName.Function} folder.`;
export const backendStartStartMessageNext = `executing 'npm run dev:teamsfx' under ${FolderName.Function} folder.`;
export const backendWatchStartMessage = `executing 'tsc --watch' under ${FolderName.Function} folder.`;
export const backendWatchStartMessageNext = `executing 'npm run watch:teamsfx' under ${FolderName.Function} folder.`;
export const botInstallStartMessage = `executing 'npm install' under ${FolderName.Bot} folder.`;
export const botStartStartMessage = "starting bot.";
export const botStartStartMessageNext = `executing 'npm run dev:teamsfx' under ${FolderName.Bot} folder.`;
export const botWatchStartMessage = `executing 'npm run watch:teamsfx' under ${FolderName.Bot} folder.`;
export const botAzuriteStartMessage = `executing 'npm run prepare-storage:teamsfx' under ${FolderName.Bot} folder.`;
export const ngrokStartStartMessage = `executing 'ngrok http' under ${FolderName.Bot} folder.`;

export const previewTitle = "preview";
export const previewStartMessage = "opening Teams web client.";
export const previewSPFxTitle = "spfx preview";
export const previewSPFxStartMessage = "opening SharePoint workbench.";
export const previewTeamsDesktopClientMessage = "opening Teams desktop client.";

export const frontendLocalEnvPrefix = "FRONTEND_";
export const backendLocalEnvPrefix = "BACKEND_";
export const authLocalEnvPrefix = "AUTH_";
export const authServicePathEnvKey = "AUTH_SERVICE_PATH";
export const botLocalEnvPrefix = "BOT_";

export const doctorResult = {
  NodeNotFound: `Cannot find Node.js.`,
  NodeNotSupported: `Node.js (@CurrentVersion) is not the officially supported version (@SupportedVersions). Your project may continue to work but we recommend to install the supported version.`,
  NodeSuccess: `Node.js version (@Version) is installed`,
  InstallNode: "Go to https://nodejs.org to install LTS Node.js.",
  InstallNodeV3:
    "The supported node versions are specified in the package.json. Go to https://nodejs.org to install a supported Node.js.",
  SideLoadingDisabled:
    "Your Microsoft 365 tenant admin hasn't enabled custom app upload permission for your account. You can't install your app to Microsoft 365!",
  NotSignIn: "No Microsoft 365 account login",
  SignInSuccess: `Microsoft 365 Account (@account) is signed in and custom app upload permission is enabled`,
  SkipTrustingCert: "Skip trusting development certificate for localhost",
  HelpLink: `Please refer to @Link for more information.`,
  NgrokWarning:
    "This software downloads npm package ngrok@4.3.3 which contains NGROK(https://ngrok.com/) v2.3.40. Customer must have a valid license to use NGROK software. Microsoft does not license use of the NGROK.",
};

export const runCommand = {
  detectRunCommand: "Option 'run-command' is not provided, set to: ",
  showCommand: "Executing command - ",
  showRunningPattern: "Running pattern - ",
  showWorkingFolder: "Working folder - ",
};

export const manifestChangesHintMessage = (): string =>
  `If you changed the manifest file, please run '${process.env.TEAMSFX_CLI_BIN_NAME} provision %s' to install app again.`;
export const m365TenantHintMessage =
  "WARN: Please note that after you enrolled your developer tenant in Office 365 Target Release, it may take a couple of days for the enrollment to take effect. Please click https://aka.ms/teamsfx-m365-apps-prerequisites for more information about setting up dev environment for extending apps across Microsoft 365.";
export const m365SwitchedMessage =
  "WARN: You are now using a different Microsoft 365 tenant from what you previously used. Please visit https://aka.ms/teamsfx-switch-tenant to get more info.";

export const defaultExecPath = "devTools/func";
