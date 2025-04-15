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
export { MicrosoftTeamsV1D0 } from "./MicrosoftTeams.v1d0";
export { MicrosoftTeamsV1D1 } from "./MicrosoftTeams.v1d1";
export { MicrosoftTeamsV1D2 } from "./MicrosoftTeams.v1d2";
export { MicrosoftTeamsV1D3 } from "./MicrosoftTeams.v1d3";
export { MicrosoftTeamsV1D4 } from "./MicrosoftTeams.v1d4";
export { MicrosoftTeamsV1D5 } from "./MicrosoftTeams.v1d5";
export { MicrosoftTeamsV1D6 } from "./MicrosoftTeams.v1d6";
export { MicrosoftTeamsV1D7 } from "./MicrosoftTeams.v1d7";
export { MicrosoftTeamsV1D8 } from "./MicrosoftTeams.v1d8";
export { MicrosoftTeamsV1D9 } from "./MicrosoftTeams.v1d9";
export { MicrosoftTeamsV1D10 } from "./MicrosoftTeams.v1d10";
export { MicrosoftTeamsV1D11 } from "./MicrosoftTeams.v1d11";
export { MicrosoftTeamsV1D12 } from "./MicrosoftTeams.v1d12";
export { MicrosoftTeamsV1D13 } from "./MicrosoftTeams.v1d13";
export { MicrosoftTeamsV1D14 } from "./MicrosoftTeams.v1d14";
export { MicrosoftTeamsV1D15 } from "./MicrosoftTeams.v1d15";
export { MicrosoftTeamsV1D16 } from "./MicrosoftTeams.v1d16";
export { MicrosoftTeamsV1D17 } from "./MicrosoftTeams.v1d17";
export { MicrosoftTeamsV1D19 } from "./MicrosoftTeams.v1d19";
export { MicrosoftTeamsV1D20 } from "./MicrosoftTeams.v1d20";
export { MicrosoftTeamsVDevPreview } from "./MicrosoftTeams.vDevPreview";
export { MicrosoftTeamsVDevPreview as DevPreviewSchema } from "./MicrosoftTeams.vDevPreview";
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
const versionToConvert = {
  "1.0": MicrosoftTeamsV1D0Convert.toMicrosoftTeamsV1D0,
  "1.1": MicrosoftTeamsV1D1Convert.toMicrosoftTeamsV1D1,
  "1.2": MicrosoftTeamsV1D2Convert.toMicrosoftTeamsV1D2,
  "1.3": MicrosoftTeamsV1D3Convert.toMicrosoftTeamsV1D3,
  "1.4": MicrosoftTeamsV1D4Convert.toMicrosoftTeamsV1D4,
  "1.5": MicrosoftTeamsV1D5Convert.toMicrosoftTeamsV1D5,
  "1.6": MicrosoftTeamsV1D6Convert.toMicrosoftTeamsV1D6,
  "1.7": MicrosoftTeamsV1D7Convert.toMicrosoftTeamsV1D7,
  "1.8": MicrosoftTeamsV1D8Convert.toMicrosoftTeamsV1D8,
  "1.9": MicrosoftTeamsV1D9Convert.toMicrosoftTeamsV1D9,
  "1.10": MicrosoftTeamsV1D10Convert.toMicrosoftTeamsV1D10,
  "1.11": MicrosoftTeamsV1D11Convert.toMicrosoftTeamsV1D11,
  "1.12": MicrosoftTeamsV1D12Convert.toMicrosoftTeamsV1D12,
  "1.13": MicrosoftTeamsV1D13Convert.toMicrosoftTeamsV1D13,
  "1.14": MicrosoftTeamsV1D14Convert.toMicrosoftTeamsV1D14,
  "1.15": MicrosoftTeamsV1D15Convert.toMicrosoftTeamsV1D15,
  "1.16": MicrosoftTeamsV1D16Convert.toMicrosoftTeamsV1D16,
  "1.17": MicrosoftTeamsV1D17Convert.toMicrosoftTeamsV1D17,
  "1.19": MicrosoftTeamsV1D19Convert.toMicrosoftTeamsV1D19,
  "1.20": MicrosoftTeamsV1D20Convert.toMicrosoftTeamsV1D20,
  devPreview: MicrosoftTeamsVDevPreviewConvert.toMicrosoftTeamsVDevPreview,
};

export function jsonToManifest(json: string): MicrosoftTeamsManifest {
  const parsed = JSON.parse(json);
  const manifestVersion = parsed.manifestVersion as string;
  const converter = versionToConvert[manifestVersion as keyof typeof versionToConvert];
  if (converter) {
    return converter(json);
  }
  throw new Error(`Unsupported manifest version: ${manifestVersion}`);
}
