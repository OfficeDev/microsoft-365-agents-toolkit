enum ATKConfigComponentName {
  AppManifest = "app-manifest",
  Basic = "atk-basic",
  VSCodeLocalDebug = "vscode-local-debug",
  Playground = "playground",
  TeamsTab = "teams-tab",
  TeamsBot = "teams-bot",
  DA = "declarative-agent",
}

enum ATKFeatureName {
  DevTunnel = "dev-tunnel",
}

export const templateConfigs = [
  {
    id: "basic-tab-ts",
    components: [
      {
        name: ATKConfigComponentName.AppManifest,
      },
      {
        name: ATKConfigComponentName.Basic,
      },
      {
        name: ATKConfigComponentName.VSCodeLocalDebug,
      },
      {
        name: ATKConfigComponentName.TeamsTab,
      },
    ],
    features: {
      [ATKFeatureName.DevTunnel]: false,
    },
  },
  {
    id: "custom-copilot-rag-azure-ai-search-ts",
    components: [
      {
        name: ATKConfigComponentName.AppManifest,
      },
      {
        name: ATKConfigComponentName.Basic,
      },
      {
        name: ATKConfigComponentName.VSCodeLocalDebug,
      },
      {
        name: ATKConfigComponentName.Playground,
      },
      {
        name: ATKConfigComponentName.TeamsBot,
      },
    ],
    features: {
      [ATKFeatureName.DevTunnel]: true,
    },
  },
];
