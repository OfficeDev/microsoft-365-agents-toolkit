// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
"use strict";

export class FolderName {
  static readonly Frontend = "tabs";
  static readonly Bot = "bot";
  static readonly Function = "api";
  static readonly SPFx = "SPFx";
  static readonly VideoFilter = "app";
}

export const baseNpmInstallCommand = "npm install";
export const defaultNpmInstallArg = "--no-audit";
export const npmInstallCommand = `${baseNpmInstallCommand} ${defaultNpmInstallArg}`;

export const LocalEnvAuthKeys = Object.freeze({
  ClientId: "AUTH_CLIENT_ID",
  ClientSecret: "AUTH_CLIENT_SECRET",
  IdentifierUri: "AUTH_IDENTIFIER_URI",
  AadMetadataAddress: "AUTH_AAD_METADATA_ADDRESS",
  OauthAuthority: "AUTH_OAUTH_AUTHORITY",
  TabEndpoint: "AUTH_TAB_APP_ENDPOINT",
  AllowedAppIds: "AUTH_ALLOWED_APP_IDS",
  Urls: "AUTH_urls",
  ServicePath: "AUTH_SERVICE_PATH",
});

export const LocalEnvBackendKeys = Object.freeze({
  WebJobsStorage: "BACKEND_AzureWebJobsStorage",
  FuncWorkerRuntime: "BACKEND_FUNCTIONS_WORKER_RUNTIME",
  AuthorityHost: "BACKEND_M365_AUTHORITY_HOST",
  TenantId: "BACKEND_M365_TENANT_ID",
  ClientId: "BACKEND_M365_CLIENT_ID",
  ClientSecret: "BACKEND_M365_CLIENT_SECRET",
  SqlEndpoint: "BACKEND_SQL_ENDPOINT",
  SqlDbName: "BACKEND_SQL_DATABASE_NAME",
  SqlUserName: "BACKEND_SQL_USER_NAME",
  SqlPassword: "BACKEND_SQL_PASSWORD",
  IdentityId: "BACKEND_IDENTITY_ID",
  ApiEndpoint: "BACKEND_API_ENDPOINT",
  ApplicationIdUri: "BACKEND_M365_APPLICATION_ID_URI",
  AllowedAppIds: "BACKEND_ALLOWED_APP_IDS",
});

export const LocalEnvBotKeys = Object.freeze({
  BotId: "BOT_BOT_ID",
  BotPassword: "BOT_BOT_PASSWORD",
  ClientId: "BOT_M365_CLIENT_ID",
  ClientSecret: "BOT_M365_CLIENT_SECRET",
  TenantID: "BOT_M365_TENANT_ID",
  OauthAuthority: "BOT_M365_AUTHORITY_HOST",
  LoginEndpoint: "BOT_INITIATE_LOGIN_ENDPOINT",
  SqlEndpoint: "BOT_SQL_ENDPOINT",
  SqlDbName: "BOT_SQL_DATABASE_NAME",
  SqlUserName: "BOT_SQL_USER_NAME",
  SqlPassword: "BOT_SQL_PASSWORD",
  IdentityId: "BOT_IDENTITY_ID",
  ApiEndpoint: "BOT_API_ENDPOINT",
  ApplicationIdUri: "BOT_M365_APPLICATION_ID_URI",
});

export const LocalEnvCertKeys = Object.freeze({
  SslCrtFile: "FRONTEND_SSL_CRT_FILE",
  SslKeyFile: "FRONTEND_SSL_KEY_FILE",
});

export const LocalEnvFrontendKeys = Object.freeze({
  Browser: "FRONTEND_BROWSER",
  Https: "FRONTEND_HTTPS",
  Port: "FRONTEND_PORT",
  TeamsFxEndpoint: "FRONTEND_REACT_APP_TEAMSFX_ENDPOINT",
  LoginUrl: "FRONTEND_REACT_APP_START_LOGIN_PAGE_URL",
  FuncEndpoint: "FRONTEND_REACT_APP_FUNC_ENDPOINT",
  FuncName: "FRONTEND_REACT_APP_FUNC_NAME",
  ClientId: "FRONTEND_REACT_APP_CLIENT_ID",
});

export class LocalDebugCertificate {
  public static readonly CertFileName: string = "localhost.crt";
  public static readonly KeyFileName: string = "localhost.key";
  public static readonly FriendlyName: string = "TeamsFx Development Certificate";
}

export const BotHostTypeName = "host-type";
export const BotHostTypes = Object.freeze({
  AppService: "app-service",
  AzureFunctions: "azure-functions",
});

export const BotCapabilities = "capabilities";

export const TaskCommand = Object.freeze({
  checkPrerequisites: "debug-check-prerequisites",
  npmInstall: "debug-npm-install",
  startLocalTunnel: "debug-start-local-tunnel",
  setUpTab: "debug-set-up-tab",
  setUpBot: "debug-set-up-bot",
  setUpSSO: "debug-set-up-sso",
  prepareManifest: "debug-prepare-manifest",
  launchWebClient: "launch-web-client",
  provision: "provision",
  deploy: "deploy",
  migrate: "migrate",
  launchDesktopClient: "launch-desktop-client",
});

export const TeamsFxNpmCommands = Object.freeze({
  startApplication: "npm run dev:teamsfx",
  startApplicationForTestTool: "npm run dev:teamsfx:testtool",
  startTestTool: "npm run dev:teamsfx:launch-testtool",
});

export const TaskOverallLabel = Object.freeze({
  NextDefault: "Pre Debug Check & Start All",
  NextM365: "Pre Debug Check & Start All & Install App",
  NextSPFx: "prepare dev env",
  TransparentDefault: "Start Microsoft 365 App Locally",
  TransparentM365: "Start Microsoft 365 App Locally & Install App",
  TestToolDefault: "Start Microsoft 365 App (Test Tool)",
});

export const TaskLabel = Object.freeze({
  PrerequisiteCheck: "Validate & install prerequisites",
  PrerequisiteCheckV3: "Validate prerequisites",
  PrerequisiteCheckV3TestTool: "Validate prerequisites (Test Tool)",
  InstallNpmPackages: "Install npm packages",
  StartLocalTunnel: "Start local tunnel",
  SetUpTab: "Set up tab",
  SetUpBot: "Set up bot",
  SetUpSSO: "Set up SSO",
  PrepareManifest: "Build & upload Teams manifest",
  InstallAzureFuncBindingExt: "Install Azure Functions binding extensions",
  StartServices: "Start services",
  StartApplication: "Start application", // V3
  StartApplicationTestTool: "Start application for Test Tool", // V3
  StartTestTool: "Start Test Tool", // V3
  StartFrontend: "Start frontend",
  StartBackend: "Start backend",
  WatchBackend: "Watch backend",
  WatchBot: "Watch bot",
  StartBot: "Start bot",
  StartAzuriteEmulator: "Start Azurite emulator",
  InstallAppInTeams: "Install app in Teams",
  GulpTrustDevCert: "gulp trust-dev-cert",
  GulpServe: "gulp serve",
  Provision: "Provision", // V3
  ProvisionSandbox: "Provision (Sandbox)",
  Deploy: "Deploy", // V3
  DeployTestTool: "Deploy (Test Tool)", // V3
  DeploySandbox: "Deploy (Sandbox)",
});

export const TaskDefaultValue = Object.freeze({
  checkPrerequisites: {
    ports: {
      tabService: 53000,
      backendService: 7071,
      backendDebug: 9229,
      botService: 3978,
      botDebug: 9239,
      spfxService: 4321,
    },
  },
  npmInstall: {
    npmInstallArgs: ["--no-audit"],
  },
  startLocalTunnel: {
    ngrokArgs: "http 3978 --log=stdout --log-format=logfmt",
    ngrokPath: "ngrok",
    writeToEnvironmentFile: {
      endpoint: "BOT_ENDPOINT",
      domain: "BOT_DOMAIN",
    },
    devTunnel: {
      bot: {
        port: 3978,
        protocol: "http",
        access: "public",
      },
    },
  },
  setUpTab: {
    baseUrl: "https://localhost:53000",
  },
  setUpBot: {
    botMessagingEndpoint: "/api/messages",
  },
  env: "local",
});

export const Prerequisite = Object.freeze({
  nodejs: "nodejs",
  m365Account: "m365Account",
  copilotAccess: "copilotAccess",
  devCert: "devCert",
  func: "func",
  ngrok: "ngrok",
  dotnet: "dotnet",
  portOccupancy: "portOccupancy",
  sandbox: "sandbox",
  vxTestApp: "vxTestApp", // TODO(aochengwang): maybe change app name
});

export const TunnelType = Object.freeze({
  devTunnel: "dev-tunnel",
  ngrok: "ngrok",
});
