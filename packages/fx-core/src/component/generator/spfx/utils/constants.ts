// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { getLocalizedString } from "../../../../common/localizeUtils";

// Licensed under the MIT license.
export class Constants {
  public static readonly PLUGIN_NAME = "SPFx";
  public static readonly PLUGIN_DEV_NAME = "fx-resource-spfx";
  public static readonly SetUpDevEnvironmentHelpLink =
    "https://aka.ms/teamsfx-spfx-dev-environment-setup";
  public static readonly TEMPLATE_NAME = "spfx-tab";
  public static readonly LatestVersion = "latest";
  public static readonly RecommendedLowestSpfxVersion = "v1.14.0";
  public static readonly GeneratorPackageName = "@microsoft/generator-sharepoint";
  public static readonly YeomanPackageName = "yo";
  public static readonly DEFAULT_WEBPART_NAME = "helloworld";
  public static readonly ScaffoldHelpLink = "https://aka.ms/teamsfx-spfx-help-v5";
  public static readonly AddWebpartHelpLink = "https://aka.ms/teamsfx-spfx-help-v5";
  public static readonly DevProgramLink =
    "https://developer.microsoft.com/en-us/microsoft-365/dev-program";
  public static readonly YO_RC_SOLUTION_NAME = "solutionName";
  public static readonly IMPORT_HELP_LINK = "https://aka.ms/teamsfx-spfx-help-v5";
  public static readonly TEAMS_APP_NAME_MAX_LENGTH = 30;
  public static readonly YO_RC_VERSION = "version";
  public static readonly YO_RC_FILE = ".yo-rc.json";
  public static readonly DEFAULT_NODE_VERSION = "18";
  public static readonly PACKAGE_JSON_FILE = "package.json";
}

export class ProgressTitleMessage {
  static readonly PreDeployProgressTitle = getLocalizedString(
    "plugins.spfx.buildSharepointPackage"
  );
  static readonly DeployProgressTitle = getLocalizedString("plugins.spfx.deploy.title");
  static readonly ScaffoldProgressTitle = getLocalizedString("plugins.spfx.scaffold.title");
  static readonly AddProgressTitle = getLocalizedString("driver.spfx.add.progress.title");
}

export class ScaffoldProgressMessage {
  static readonly DependencyCheck = getLocalizedString("plugins.spfx.scaffold.dependencyCheck");
  static readonly DependencyInstall = getLocalizedString("plugins.spfx.scaffold.dependencyInstall");
  static readonly ScaffoldProject = getLocalizedString("plugins.spfx.scaffold.scaffoldProject");
  static readonly UpdateManifest = getLocalizedString("plugins.spfx.scaffold.updateManifest");
}

export class ManifestTemplate {
  static readonly LOCAL_CONTENT_URL =
    "https://{teamSiteDomain}/_layouts/15/TeamsLogon.aspx?SPFX=true&dest=/_layouts/15/TeamsWorkBench.aspx%3Fteams%26personal%26componentId=%s%26forceLocale={locale}%26loadSPFX%3Dtrue%26debugManifestsFile%3Dhttps%3A%2F%2Flocalhost%3A4321%2Ftemp%2Fmanifests.js";
  static readonly LOCAL_CONFIGURATION_URL =
    "https://{teamSiteDomain}{teamSitePath}/_layouts/15/TeamsLogon.aspx?SPFX=true&dest={teamSitePath}/_layouts/15/TeamsWorkBench.aspx%3FcomponentId=%s%26openPropertyPane=true%26teams%26forceLocale={locale}%26loadSPFX%3Dtrue%26debugManifestsFile%3Dhttps%3A%2F%2Flocalhost%3A4321%2Ftemp%2Fmanifests.js";
  static readonly REMOTE_CONTENT_URL =
    "https://{teamSiteDomain}/_layouts/15/TeamsLogon.aspx?SPFX=true&dest=/_layouts/15/teamshostedapp.aspx%3Fteams%26personal%26componentId=%s%26forceLocale={locale}";
  static readonly REMOTE_CONFIGURATION_URL =
    "{{^config.isLocalDebug}}https://{teamSiteDomain}{teamSitePath}/_layouts/15/TeamsLogon.aspx?SPFX=true&dest={teamSitePath}/_layouts/15/teamshostedapp.aspx%3FopenPropertyPane=true%26teams%26componentId=%s%26forceLocale={locale}{{/config.isLocalDebug}}{{#config.isLocalDebug}}https://{teamSiteDomain}{teamSitePath}/_layouts/15/TeamsLogon.aspx?SPFX=true&dest={teamSitePath}/_layouts/15/TeamsWorkBench.aspx%3FcomponentId=%s%26openPropertyPane=true%26teams%26forceLocale={locale}%26loadSPFX%3Dtrue%26debugManifestsFile%3Dhttps%3A%2F%2Flocalhost%3A4321%2Ftemp%2Fmanifests.js{{/config.isLocalDebug}}";
  static readonly WEBSITE_URL = "https://products.office.com/en-us/sharepoint/collaboration";
  static readonly WEB_APP_INFO_RESOURCE = "https://{teamSiteDomain}";
  static readonly WEB_APP_INFO_ID = "00000003-0000-0ff1-ce00-000000000000";
}
