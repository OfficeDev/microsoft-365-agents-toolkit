// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

export class Constants {
  static readonly ActionName = "spfx/add";

  static readonly LOCAL_CONTENT_URL =
    "https://{teamSiteDomain}/_layouts/15/TeamsLogon.aspx?SPFX=true&dest=/_layouts/15/TeamsWorkBench.aspx%3Fteams%26personal%26componentId=%s%26forceLocale={locale}%26loadSPFX%3Dtrue%26debugManifestsFile%3Dhttps%3A%2F%2Flocalhost%3A4321%2Ftemp%2Fmanifests.js";
  static readonly REMOTE_CONTENT_URL =
    "https://{teamSiteDomain}/_layouts/15/TeamsLogon.aspx?SPFX=true&dest=/_layouts/15/teamshostedapp.aspx%3Fteams%26personal%26componentId=%s%26forceLocale={locale}";
  static readonly YO_RC_FILE = ".yo-rc.json";
}
