// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import { MicrosoftTeamsV1D0, Convert as MicrosoftTeamsV1D0Convert } from "./MicrosoftTeams.v1d0";
import { MicrosoftTeamsV1D1, Convert as MicrosoftTeamsV1D1Convert } from "./MicrosoftTeams.v1d1";
import { MicrosoftTeamsV1D2, Convert as MicrosoftTeamsV1D2Convert } from "./MicrosoftTeams.v1d2";
import { MicrosoftTeamsV1D3, Convert as MicrosoftTeamsV1D3Convert } from "./MicrosoftTeams.v1d3";
import { MicrosoftTeamsV1D4, Convert as MicrosoftTeamsV1D4Convert } from "./MicrosoftTeams.v1d4";
import { MicrosoftTeamsV1D5, Convert as MicrosoftTeamsV1D5Convert } from "./MicrosoftTeams.v1d5";
import { MicrosoftTeamsV1D6, Convert as MicrosoftTeamsV1D6Convert } from "./MicrosoftTeams.v1d6";
import { MicrosoftTeamsV1D7, Convert as MicrosoftTeamsV1D7Convert } from "./MicrosoftTeams.v1d7";
import { MicrosoftTeamsV1D8, Convert as MicrosoftTeamsV1D8Convert } from "./MicrosoftTeams.v1d8";
import { MicrosoftTeamsV1D9, Convert as MicrosoftTeamsV1D9Convert } from "./MicrosoftTeams.v1d9";
import { MicrosoftTeamsV1D10, Convert as MicrosoftTeamsV1D10Convert } from "./MicrosoftTeams.v1d10";
import { MicrosoftTeamsV1D11, Convert as MicrosoftTeamsV1D11Convert } from "./MicrosoftTeams.v1d11";
import { MicrosoftTeamsV1D12, Convert as MicrosoftTeamsV1D12Convert } from "./MicrosoftTeams.v1d12";
import { MicrosoftTeamsV1D13, Convert as MicrosoftTeamsV1D13Convert } from "./MicrosoftTeams.v1d13";
import { MicrosoftTeamsV1D14, Convert as MicrosoftTeamsV1D14Convert } from "./MicrosoftTeams.v1d14";
import { MicrosoftTeamsV1D15, Convert as MicrosoftTeamsV1D15Convert } from "./MicrosoftTeams.v1d15";
import { MicrosoftTeamsV1D16, Convert as MicrosoftTeamsV1D16Convert } from "./MicrosoftTeams.v1d16";
import { MicrosoftTeamsV1D17, Convert as MicrosoftTeamsV1D17Convert } from "./MicrosoftTeams.v1d17";
import { MicrosoftTeamsV1D19, Convert as MicrosoftTeamsV1D19Convert } from "./MicrosoftTeams.v1d19";
import { MicrosoftTeamsV1D20, Convert as MicrosoftTeamsV1D20Convert } from "./MicrosoftTeams.v1d20";
import {
  MicrosoftTeamsVDevPreview,
  Convert as MicrosoftTeamsVDevPreviewConvert,
} from "./MicrosoftTeams.vDevPreview";

export type MicrosoftTeamsManifest =
  | MicrosoftTeamsV1D0
  | MicrosoftTeamsV1D1
  | MicrosoftTeamsV1D2
  | MicrosoftTeamsV1D3
  | MicrosoftTeamsV1D4
  | MicrosoftTeamsV1D5
  | MicrosoftTeamsV1D6
  | MicrosoftTeamsV1D7
  | MicrosoftTeamsV1D8
  | MicrosoftTeamsV1D9
  | MicrosoftTeamsV1D10
  | MicrosoftTeamsV1D11
  | MicrosoftTeamsV1D12
  | MicrosoftTeamsV1D13
  | MicrosoftTeamsV1D14
  | MicrosoftTeamsV1D15
  | MicrosoftTeamsV1D16
  | MicrosoftTeamsV1D17
  | MicrosoftTeamsV1D19
  | MicrosoftTeamsV1D20
  | MicrosoftTeamsVDevPreview;

export function jsonToManifest(json: string): MicrosoftTeamsManifest {
  const parsed = JSON.parse(json);
  const manifestVersion = parsed.manifestVersion as string;
  if (manifestVersion === "1.0") {
    return MicrosoftTeamsV1D0Convert.toMicrosoftTeamsV1D0(json);
  } else if (manifestVersion === "1.1") {
    return MicrosoftTeamsV1D1Convert.toMicrosoftTeamsV1D1(json);
  } else if (manifestVersion === "1.2") {
    return MicrosoftTeamsV1D2Convert.toMicrosoftTeamsV1D2(json);
  } else if (manifestVersion === "1.3") {
    return MicrosoftTeamsV1D3Convert.toMicrosoftTeamsV1D3(json);
  } else if (manifestVersion === "1.4") {
    return MicrosoftTeamsV1D4Convert.toMicrosoftTeamsV1D4(json);
  } else if (manifestVersion === "1.5") {
    return MicrosoftTeamsV1D5Convert.toMicrosoftTeamsV1D5(json);
  } else if (manifestVersion === "1.6") {
    return MicrosoftTeamsV1D6Convert.toMicrosoftTeamsV1D6(json);
  } else if (manifestVersion === "1.7") {
    return MicrosoftTeamsV1D7Convert.toMicrosoftTeamsV1D7(json);
  } else if (manifestVersion === "1.8") {
    return MicrosoftTeamsV1D8Convert.toMicrosoftTeamsV1D8(json);
  } else if (manifestVersion === "1.9") {
    return MicrosoftTeamsV1D9Convert.toMicrosoftTeamsV1D9(json);
  } else if (manifestVersion === "1.10") {
    return MicrosoftTeamsV1D10Convert.toMicrosoftTeamsV1D10(json);
  } else if (manifestVersion === "1.11") {
    return MicrosoftTeamsV1D11Convert.toMicrosoftTeamsV1D11(json);
  } else if (manifestVersion === "1.12") {
    return MicrosoftTeamsV1D12Convert.toMicrosoftTeamsV1D12(json);
  } else if (manifestVersion === "1.13") {
    return MicrosoftTeamsV1D13Convert.toMicrosoftTeamsV1D13(json);
  } else if (manifestVersion === "1.14") {
    return MicrosoftTeamsV1D14Convert.toMicrosoftTeamsV1D14(json);
  } else if (manifestVersion === "1.15") {
    return MicrosoftTeamsV1D15Convert.toMicrosoftTeamsV1D15(json);
  } else if (manifestVersion === "1.16") {
    return MicrosoftTeamsV1D16Convert.toMicrosoftTeamsV1D16(json);
  } else if (manifestVersion === "1.17") {
    return MicrosoftTeamsV1D17Convert.toMicrosoftTeamsV1D17(json);
  } else if (manifestVersion === "1.19") {
    return MicrosoftTeamsV1D19Convert.toMicrosoftTeamsV1D19(json);
  } else if (manifestVersion === "1.20") {
    return MicrosoftTeamsV1D20Convert.toMicrosoftTeamsV1D20(json);
  } else if (manifestVersion === "devPreview") {
    return MicrosoftTeamsVDevPreviewConvert.toMicrosoftTeamsVDevPreview(json);
  }
  throw new Error(`Unsupported manifest version: ${manifestVersion}`);
}
const json = {
  $schema:
    "https://developer.microsoft.com/en-us/json-schemas/teams/v1.19/MicrosoftTeams.schema.json",
  manifestVersion: "1.19",
  version: 1.0,
  id: "${{TEAMS_APP_ID}}",
  developer: {
    name: "Teams App, Inc.",
    websiteUrl: "https://www.example.com",
    privacyUrl: "https://www.example.com/privacy",
    termsOfUseUrl: "https://www.example.com/termofuse",
  },
  icons: {
    color: "color.png",
    outline: "outline.png",
  },
  name: {
    short: "huajiecea040906${{APP_NAME_SUFFIX}}",
    full: "full name for huajiecea040906",
  },
  description: {
    short: "Repair Service",
    full: "A simple service to manage repairs",
  },
  accentColor: "#FFFFFF",
  bots: [
    {
      botId: "${{BOT_ID}}",
      scopes: ["personal", "team", "groupChat"],
      supportsFiles: false,
      isNotificationOnly: false,
      commandLists: [
        {
          scopes: ["personal"],
          commands: [
            {
              title: "List all repairs without auth",
              description: "List all repairs without auth",
            },
          ],
        },
      ],
    },
  ],
  composeExtensions: [],
  configurableTabs: [],
  staticTabs: [],
  permissions: ["identity", "messageTeamMembers"],
  validDomains: [],
};
console.log(jsonToManifest(JSON.stringify(json)));

// ManifestUtil.validateManifest(json as Manifest).then((result) => {
//   console.log(result);
// });
