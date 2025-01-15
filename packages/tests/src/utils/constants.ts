// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import * as os from "os";

export class Extension {
  public static readonly displayName: string = "Teams Toolkit";
  public static readonly treeViewSectionName: string = "Development";
  public static readonly activatedItemName: string = "DEVELOPMENT";
  public static readonly sidebarWelcomeSectionName: string = "Teams Toolkit";
  public static readonly sidebarWelcomeContentName: string = "Create a New App";
  public static readonly sidebarCommandContentName: string = "Create New App";
  public static readonly sidebarCommandContentNameOfficeDev: string =
    "Preview Your Office Add-in (F5)";
  public static readonly settingsCategory: string = "Fx-extension";
  public static readonly settingsInsiderPreview: string = "Insider Preview";
}

export class Project {
  public static readonly namePrefix = "fxui";
}

export class TeamsFxProject {
  public static readonly ExtensionConfigFolderName = "fx";
  public static readonly TelemetryLoggerFileName = "telemetryTest.log";
  public static readonly TelemetryLoggerFilePath =
    os.homedir +
    `/.${TeamsFxProject.ExtensionConfigFolderName}/${TeamsFxProject.TelemetryLoggerFileName}`;
}

export enum TemplateProject {
  HelloWorldTabBackEnd = "Tab App with Azure Backend",
  ContactExporter = "Contact Exporter using Graph Toolkit",
  OneProductivityHub = "One Productivity Hub using Graph Toolkit",
  HelloWorldBotSSO = "Bot App with SSO Enabled",
  TodoListBackend = "Todo List with backend on Azure",
  TodoListSpfx = "Todo List with SPFx",
  ShareNow = "Share Now",
  MyFirstMeeting = "My First Meeting App",
  TodoListM365 = "Todo List (Works in Teams, Outlook and Office)",
  NpmSearch = "NPM Search Connector",
  ProactiveMessaging = "Proactive Messaging",
  AdaptiveCard = "Adaptive Card Notification",
  IncomingWebhook = "Incoming Webhook Notification",
  GraphConnector = "Graph Connector App",
  StockUpdate = "Stocks Update",
  QueryOrg = "Org User Search Connector",
  Deeplinking = "Hello World Deeplinking Navigation Tab App",
  Dashboard = "Team Central Dashboard",
  AssistDashboard = "Developer Assist Dashboard",
  DiceRoller = "Dice Roller in meeting",
  OutlookTab = "Hello World Teams Tab and Outlook add-in",
  OutlookSignature = "Set signature using Outlook add-in",
  ChefBot = "Teams Chef Bot",
  GraphConnectorBot = "Graph Connector Bot",
  SpfxProductivity = "One Productivity Hub using Graph Toolkit with SPFx",
  RetailDashboard = "Contoso Retail Dashboard",
  TabSSOApimProxy = "SSO Enabled Tab via APIM Proxy",
  LargeScaleBot = "Large Scale Notification Bot",
  BotSSODocker = "Containerized Bot App with SSO Enabled",
  HelloWorldTabDocker = "Containerized Hello World Tab with Backend",
  FoodCatalog = "Ingest Custom API Data into Microsoft 365 with a Microsoft Graph Connector",
  RedditLink = "Format Reddit Link into Adaptive Card",
  IntelligentDataChart = "Intelligent Data Chart Generator",
}

export enum TemplateProjectFolder {
  HelloWorldTabBackEnd = "hello-world-tab-with-backend",
  ContactExporter = "graph-toolkit-contact-exporter",
  HelloWorldBotSSO = "bot-sso",
  BotSSODocker = "bot-sso-docker",
  TabDocker = "hello-world-tab-docker",
  TodoListSpfx = "todo-list-SPFx",
  MyFirstMeeting = "hello-world-in-meeting",
  TodoListM365 = "todo-list-with-Azure-backend-M365",
  NpmSearch = "NPM-search-connector-M365",
  ProactiveMessaging = "bot-proactive-messaging-teamsfx",
  AdaptiveCard = "adaptive-card-notification",
  IncomingWebhook = "incoming-webhook-notification",
  StockUpdate = "stocks-update-notification-bot",
  QueryOrg = "query-org-user-with-message-extension-sso",
  GraphConnector = "graph-connector-app",
  OneProductivityHub = "graph-toolkit-one-productivity-hub",
  TodoListBackend = "todo-list-with-Azure-backend",
  ShareNow = "share-now",
  // v3 only
  Dashboard = "team-central-dashboard",
  OutlookSignature = "outlook-set-signature",
  OutlookTab = "hello-world-teams-tab-and-outlook-add-in",
  AssistDashboard = "developer-assist-dashboard",
  DiceRoller = "live-share-dice-roller",
  ChefBot = "a.teamsChefBot",
  GraphConnectorBot = "graph-connector-bot",
  SpfxProductivity = "spfx-productivity-dashboard",
  RetailDashboard = "react-retail-dashboard",
  TabSSOApimProxy = "sso-enabled-tab-via-apim-proxy",
  LargeScaleBot = "large-scale-notification",
  HelloWorldTabDocker = "hello-world-tab-docker",
  FoodCatalog = "nodejs-typescript-food-catalog",
  RedditLink = "nodejs",
  IntelligentDataChart = "intelligent-data-chart-generator",
  // v2 only
  Deeplinking = "deep-linking-hello-world-tab-without-sso-M365",
}

export const sampleProjectMap: Record<TemplateProject, TemplateProjectFolder> =
  {
    [TemplateProject.HelloWorldTabBackEnd]:
      TemplateProjectFolder.HelloWorldTabBackEnd,
    [TemplateProject.ContactExporter]: TemplateProjectFolder.ContactExporter,
    [TemplateProject.OneProductivityHub]:
      TemplateProjectFolder.OneProductivityHub,
    [TemplateProject.HelloWorldBotSSO]: TemplateProjectFolder.HelloWorldBotSSO,
    [TemplateProject.TodoListBackend]: TemplateProjectFolder.TodoListBackend,
    [TemplateProject.TodoListSpfx]: TemplateProjectFolder.TodoListSpfx,
    [TemplateProject.ShareNow]: TemplateProjectFolder.ShareNow,
    [TemplateProject.MyFirstMeeting]: TemplateProjectFolder.MyFirstMeeting,
    [TemplateProject.TodoListM365]: TemplateProjectFolder.TodoListM365,
    [TemplateProject.NpmSearch]: TemplateProjectFolder.NpmSearch,
    [TemplateProject.ProactiveMessaging]:
      TemplateProjectFolder.ProactiveMessaging,
    [TemplateProject.AdaptiveCard]: TemplateProjectFolder.AdaptiveCard,
    [TemplateProject.IncomingWebhook]: TemplateProjectFolder.IncomingWebhook,
    [TemplateProject.GraphConnector]: TemplateProjectFolder.GraphConnector,
    [TemplateProject.StockUpdate]: TemplateProjectFolder.StockUpdate,
    [TemplateProject.QueryOrg]: TemplateProjectFolder.QueryOrg,
    [TemplateProject.Deeplinking]: TemplateProjectFolder.Deeplinking,
    [TemplateProject.Dashboard]: TemplateProjectFolder.Dashboard,
    [TemplateProject.OutlookSignature]: TemplateProjectFolder.OutlookSignature,
    [TemplateProject.OutlookTab]: TemplateProjectFolder.OutlookTab,
    [TemplateProject.AssistDashboard]: TemplateProjectFolder.AssistDashboard,
    [TemplateProject.DiceRoller]: TemplateProjectFolder.DiceRoller,
    [TemplateProject.ChefBot]: TemplateProjectFolder.ChefBot,
    [TemplateProject.GraphConnectorBot]:
      TemplateProjectFolder.GraphConnectorBot,
    [TemplateProject.SpfxProductivity]: TemplateProjectFolder.SpfxProductivity,
    [TemplateProject.RetailDashboard]: TemplateProjectFolder.RetailDashboard,
    [TemplateProject.TabSSOApimProxy]: TemplateProjectFolder.TabSSOApimProxy,
    [TemplateProject.LargeScaleBot]: TemplateProjectFolder.LargeScaleBot,
    [TemplateProject.BotSSODocker]: TemplateProjectFolder.BotSSODocker,
    [TemplateProject.HelloWorldTabDocker]:
      TemplateProjectFolder.HelloWorldTabDocker,
    [TemplateProject.FoodCatalog]: TemplateProjectFolder.FoodCatalog,
    [TemplateProject.RedditLink]: TemplateProjectFolder.RedditLink,
    [TemplateProject.IntelligentDataChart]:
      TemplateProjectFolder.IntelligentDataChart,
  };

export enum Resource {
  AzureKeyVault = "azure-keyvault",
  AzureFunction = "azure-function",
  AzureApim = "azure-apim",
  AzureSql = "azure-sql",
}

export enum ResourceToDeploy {
  Spfx = "spfx",
  FrontendHosting = "frontend-hosting",
  Bot = "bot",
  Function = "azure-function",
  Apim = "apim",
  AadManifest = "aad-manifest",
}

export enum Capability {
  Bot = "bot",
  Notification = "notification",
  CommandBot = "command-bot",
  WorkflowBot = "workflow-bot",
  TabNonSso = "tab-non-sso",
  M365SsoLaunchPage = "sso-launch-page",
  DashboardTab = "dashboard-tab",
  Spfx = "tab-spfx",
  M365SearchApp = "search-app",
  MessageExtension = "message-extension",
  LinkUnfurling = "link-unfurling",
  // v2 only
  Tab = "tab",
  // v3 only
  AiBot = "custom-copilot-basic",
  RAG = "custom-copilot-rag",
  Agent = "custom-copilot-agent",
  TaskPane = "taskpane",
  ApiPlugin = "api-plugin",
  DeclarativeAgent = "declarative-agent",
}

export enum Trigger {
  Http = "http-functions",
  Express = "http-express",
  Timmer = "timer-functions",
}

export enum Framework {
  React = "react",
  Minimal = "minimal",
  None = "none",
}

export class Timeout {
  /**
   * Wait a while to ensure the input is ready
   */
  public static readonly input: number = 500;

  /**
   * Wait until the command takes effect
   */
  public static readonly command: number = 4 * 60 * 1000;

  /**
   * Wait until the webView takes effect
   */
  public static readonly webView: number = 20 * 1000;

  /**
   * Wait for some time to take effect
   */
  public static readonly shortTimeWait: number = 5 * 1000;
  public static readonly shortTimeLoading: number = 30 * 1000;
  public static readonly longTimeWait: number = 60 * 1000;
  public static readonly stopdebugging: number = 30 * 1000;
  public static readonly startdebugging: number = 30 * 1000;

  /**
   * Wait until extension is activated
   */
  public static readonly activatingExtension: number = 3 * 60 * 1000;

  /**
   * Wait until terminal exist and contains target message
   */
  public static readonly terminal: number = 12 * 60 * 1000;

  public static readonly reloadWindow: number = 60 * 1000;

  public static readonly closeDebugWindow: number = 30 * 1000;

  public static readonly copyBotTerminal: number = 30 * 1000;

  public static readonly installWait: number = 5 * 60 * 1000;
  /**
   * playwright
   */
  public static readonly chromiumLaunchTimeout: number = 1 * 60 * 1000;
  public static readonly playwrightDefaultTimeout: number = 2 * 60 * 1000;
  public static readonly playwrightConsentPageReload: number = 3000;
  public static readonly playwrightBotConsentContinueButton: number =
    2 * 60 * 1000;
  public static readonly playwrightConsentPopupPage: number = 30 * 1000;
  public static readonly playwrightAddAppButton: number = 180 * 1000;

  // mocha
  public static readonly prepareTestCase: number = 10 * 60 * 1000;
  public static readonly finishTestCase: number = 10 * 60 * 1000;
  public static readonly testCase: number = 30 * 60 * 1000;
  public static readonly finishAzureTestCase: number = 15 * 60 * 1000;
  public static readonly testAzureCase: number = 45 * 60 * 1000;
  public static readonly migrationTestCase: number = 40 * 60 * 1000;

  // SPFx
  public static readonly spfxProvision: number = 10 * 1000;
  public static readonly spfxDeploy: number = 4 * 60 * 1000;

  // Tab
  public static readonly tabProvision: number = 5 * 60 * 1000;
  public static readonly tabDeploy: number = 6 * 60 * 1000;

  // Bot
  public static readonly botDeploy: number = 10 * 60 * 1000;

  // Add Collaborator
  public static readonly addCollaborator: number = 60 * 1000;

  // open API
  public static readonly openAPIProvision: number = 20 * 1000;
}

export class TreeViewCommands {
  public static readonly CreateProjectCommand: string = "Create New App";
  public static readonly SamplesCommand: string = "View Samples";
  public static readonly QuickStartCommand: string = "Get Started";
  public static readonly BuildTeamsPackageCommand: string =
    "Zip Teams App Package";
  public static readonly DevelopmentSectionName: string = "DEVELOPMENT";
  public static readonly DevelopmentSectionItems: string[] = [
    "Create New App",
    "View Samples",
    "View How-to Guides",
    "Preview Your Teams App (F5)",
  ];
  public static readonly EnvSectionName: string = "ENVIRONMENT";

  public static readonly OfficeDevDevelopmentSectionName: string =
    "DEVELOPMENT";
  public static readonly OfficeDevDevelopmentSectionItems: string[] = [
    "Create a New App",
    "View Samples",
    "Check and Install Dependencies",
    "Preview Your Office Add-in (F5)",
    "Stop Previewing Your Office Add-in",
  ];

  public static readonly OfficeDevLifeCycleSectionName: string = "LIFECYCLE";
  public static readonly OfficeDevLifeCycleSectionItems: string[] = [
    "Deploy",
    "Publish",
  ];

  public static readonly OfficeDevUtilitySectionName: string = "UTILITY";
  public static readonly OfficeDevUtilitySectionItems: string[] = [
    "Validate Manifest File",
    "Script Lab",
    "View Prompts for GitHub Copilot",
  ];

  public static readonly OfficeDevHelpAndFeedBackSectionName: string =
    "HELP AND FEEDBACK";
  public static readonly OfficeDevHelpAndFeedBackSectionItems: string[] = [
    "Documentation",
    "Get Started",
    "Open Partner Center",
    "Report Issues on GitHub",
  ];
}

export class CommandPaletteCommands {
  public static readonly QuickStartCommand: string = "Teams: Get Started";
  public static readonly AccountsCommand: string = "Teams: Accounts";
  public static readonly SamplesCommand: string = "Teams: View Samples";
  public static readonly CreateProjectCommand: string = "Teams: Create New App";
  public static readonly ManifestValidateCommand: string =
    "Teams: Validate manifest file";
  public static readonly BuildTeamsPackageCommand: string =
    "Teams: Zip Teams App Package";
  public static readonly ProvisionCommand: string = "Teams: Provision";
  public static readonly DeployCommand: string = "Teams: Deploy";
  public static readonly PublishCommand: string = "Teams: Publish";
  public static readonly CreateEnvironmentCommand: string =
    "Teams: Create New Environment";
  public static readonly DeployAadAppManifestCommand: string =
    "Teams: Update Microsoft Entra App";
  public static readonly UpgradeProjectCommand: string =
    "Teams: Upgrade Project";
  public static readonly InstallTTK: string =
    "Extensions: Install Specific Version of Extension";
  public static readonly AddSpfxWebPart: string = "Teams: Add SPFx web part";
}

export type AppType =
  | "tab"
  | "tabnsso"
  | "tabbot"
  | "bot"
  | "crbot" // command and response bot (name cannot be too long or it will exceed teams app name limit)
  | "funcnoti" // functions notification bot
  | "expressnoti" // express notification bot
  | "msg"
  | "msgsa"
  | "m365lp"
  | "spfx"
  | "gspfx"
  | "dashboard"
  | "workflow"
  | "timenoti"
  | "functimernoti"
  | "addin"
  | "importaddin"
  | "linkunfurl"
  | "aichat"
  | "aiagentassist"
  | "aiagentnew"
  | "chatdata"
  | "cdcustomapi" //chatadata customApi
  | "msgnewapi"
  | "msgopenapi"
  | "msgapikey"
  | "msgmicroentra"
  | "importspfx"
  | "msgmulparams"
  | "msgapikeyspec";

export class FeatureFlagName {
  static readonly InsiderPreview = "__TEAMSFX_INSIDER_PREVIEW";
}

export enum LocalDebugTaskLabel {
  StartLocalTunnel = "Start local tunnel",
  StartBot = "Start bot",
  StartBotApp = "Start application",
  StartFrontend = "Start frontend",
  StartApplication = "Start application",
  StartBackend = "Start backend",
  StartWebhook = "Start Incoming Webhook",
  WatchBackend = "Watch backend",
  InstallNpmPackages = "Install npm packages",
  ApiNpmInstall = "api npm install",
  BotNpmInstall = "bot npm install",
  TabsNpmInstall = "tabs npm install",
  SpfxNpmInstall = "SPFx npm install",
  GulpServe = "gulp serve",
  Azurite = "Start Azurite emulator",
  Compile = "Compile typescript",
  StartWebServer = "Start web server",
  DockerRun = "docker-run: debug",
  DockerTask = "docker",
  EnsureDevTunnnel = "Ensure DevTunnel",
  RunWatch = "Run watch",
  FuncStart = "func: host start",
}

export class LocalDebugTaskResult {
  static readonly FrontendSuccess = "Compiled successfully";
  static readonly FrontendReady = "ready";
  static readonly FrontendNoIssue = "webpack compiled";
  static readonly FrontendStarted = "Express server listening on";
  static readonly StartSuccess = "started successfully";
  static readonly AzuriteSuccess = "Azurite Table service is successfully";
  static readonly CompiledSuccess = "Found 0 errors";
  static readonly BotAppSuccess = "Functions:";
  static readonly AppSuccess = "Bot Started";
  static readonly GulpServeSuccess = "Running server";
  static readonly Failed = "failed";
  static readonly Error = "error";
  static readonly DebuggerAttached = "Debugger attached";
  static readonly WebServerSuccess = "press h to show help";
  static readonly DockerFinish = "press any key to close it";
  static readonly DevtunnelSuccess = "Ready to accept connections for tunnel:";
  static readonly FunctionStarted = "Worker process started and initialized";
}

export enum LocalDebugTaskLabel2 {
  StartBot2 = "Start Bot",
  PythonDebugConsole = "Python Debug Console",
  StartTestTool = "Start Test Tool",
}

export enum LocalDebugError {
  ElementNotInteractableError = "ElementNotInteractableError",
  TimeoutError = "TimeoutError",
  WarningError = "Warning",
  WarningCapError = "WARNING",
  DeprecatedError = "npm warn deprecated",
  CompiledWithWarningError = "Compiled with warnings",
}

export class LocalDebugTaskInfo {
  static readonly StartBotAppInfo = "App Started";
  static readonly StartBotInfo = "Bot Started";
  static readonly StartBotInfo2 = "Bot started";
}

export class DebugItemSelect {
  static readonly DebugInTeamsUsingChrome = "Debug in Teams (Chrome)";
  static readonly DebugInTestTool = "Debug in Test Tool";
}

export class TestFilePath {
  static readonly configurationFolder = "env";
}

export class Notification {
  static readonly Incompatible =
    "The current project is incompatible with the installed version of Teams Toolkit.";
  static readonly TaskError =
    "There are task errors. See the output for details.";
  static readonly Upgrade =
    "Upgrade your Teams Toolkit project to stay compatible with the latest version. A backup directory will be created along with an Upgrade Summary.";
  static readonly Upgrade_dicarded =
    "Please upgrade your project to stay compatible with the latest version, your current project contains configurations from an older Teams Toolkit. The auto-upgrade process will generate backups in case an error occurs.";
  static readonly ProvisionSucceeded = "successfully";
  static readonly DeploySucceeded = "successfully";
  static readonly PublishSucceeded = "successfully";
  static readonly UnresolvedPlaceholderError =
    "MissingEnvironmentVariablesError";
  static readonly ZipAppPackageSucceeded = "successfully built";
}

export class CreateProjectQuestion {
  static readonly CustomCopilot = "Custom Engine Agent";
  static readonly Bot = "Bot";
  static readonly Tab = "Tab";
  static readonly MessageExtension = "Message Extension";
  static readonly OfficeAddin = "Outlook Add-in";
  static readonly NewTeamsApp = "Start with a Teams capability";
  static readonly SpfxSharepointFrameworkInTtk = "Install the latest SPFx";
  static readonly SpfxSharepointFrameworkGlobalEnvInTtk =
    "Use globally installed SPFx";
  static readonly NewAddinApp = "Start with an Outlook add-in";
  static readonly CreateNewSpfxSolution = "Create New SPFx Solution";
  static readonly ImportExistingSpfxSolution = "Import Existing SPFx Solution";
  static readonly BuildNotificationBot = "Build a Notification Bot";
  static readonly BuildDeclarativeAgent = "Build a Declarative Agent";
}

export class ValidationContent {
  static readonly Tab = "Hello, World";
  static readonly Bot = "Your Hello World Bot is Running";
  static readonly BotWelcomeInstruction =
    "Hi there! I'm a Teams bot that will echo what you said to me";
  static readonly GraphBot = "Your Graph Connector Bot is Running";
  static readonly AiChatBotWelcomeInstruction = "How can I help you today?";
  static readonly AiAssistantBotWelcomeInstruction =
    "I'm an assistant bot. How can I help you today?";
  static readonly AiBotErrorMessage = "The bot encountered an error or bug";
  static readonly AiBotErrorMessage2 = "An AI request failed";
}

export class CliVersion {
  static readonly V2TeamsToolkitStable425 = "1.2.6";
  static readonly V2TeamsToolkit400 = "1.0.0";
}

export const ResourceGroupEnvName = "AZURE_RESOURCE_GROUP_NAME";
export const BotIdEnvName = "BOT_ID";
export const AADAppIdEnvNames = ["AAD_APP_CLIENT_ID", BotIdEnvName];
export const TeamsAppIdEnvName = "TEAMS_APP_ID";
export const M365TitleIdEnvName = "M365_TITLE_ID";

export const strings = {
  deleteResourceGroup: {
    success: `[Success] Resource group %s is deleted.`,
    failed: `[Failed] Resource group %s is not deleted.`,
    skipped: `[Skipped] Resource group %s does not exist.`,
  },
};
