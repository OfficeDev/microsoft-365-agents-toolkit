// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { AccessGithubError, SampleConfig, sendRequestWithTimeout } from "@microsoft/teamsfx-core";
import axios from "axios";

const OfficeSampleCofigOwner = "OfficeDev";
const OfficeSampleRepo = "Office-Samples";
const OfficeSampleConfigFile = ".config/samples-config-v1.json";
const OfficeSampleConfigBranch = "agent";

interface OfficeSampleCollection {
  samples: SampleConfig[];
  fileterOptions: {
    capabilities: string[];
    languages: string[];
    technologies: string[];
  };
}

type OfficeSampleConfigType = {
  samples: Array<Record<string, unknown>>;
  filterOptions: Record<string, Array<string>>;
};

class OfficeSampleProvider {
  private officeSampleCollection: OfficeSampleCollection | undefined;

  public get OfficeSampleCollection(): Promise<OfficeSampleCollection> {
    if (!this.officeSampleCollection) {
      return this.loadOfficeSamples();
    }
    return Promise.resolve(this.officeSampleCollection);
  }

  private async loadOfficeSamples(): Promise<OfficeSampleCollection> {
    const officeSamplesConfig =
      (await this.featchOfficeSamplesConfigFileContent()) as OfficeSampleConfigType;
    const officeSamples = officeSamplesConfig.samples.map((sample) => {
      return {
        ...sample,
        onboardDate: new Date(sample["onboardDate"] as string),
        downloadUrlInfo: {
          owner: OfficeSampleCofigOwner,
          repository: OfficeSampleRepo,
          ref: OfficeSampleConfigBranch,
          dir: sample["id"] as string,
        },
        gifUrl:
          sample["gifPath"] !== undefined
            ? `https://raw.githubusercontent.com/${OfficeSampleCofigOwner}/${OfficeSampleRepo}/${OfficeSampleConfigBranch}/${
                sample["id"] as string
              }/${sample["gifPath"] as string}`
            : undefined,
      } as SampleConfig;
    });
    return {
      samples: officeSamples,
      fileterOptions: {
        capabilities: officeSamplesConfig.filterOptions["capabilities"],
        languages: officeSamplesConfig.filterOptions["languages"],
        technologies: officeSamplesConfig.filterOptions["technologies"],
      },
    };
  }

  private async featchOfficeSamplesConfigFileContent(): Promise<unknown> {
    const url = `https://raw.githubusercontent.com/${OfficeSampleCofigOwner}/${OfficeSampleRepo}/${OfficeSampleConfigBranch}/${OfficeSampleConfigFile}`;
    try {
      const fileResponse = await sendRequestWithTimeout(
        async () => {
          return await axios.get(url, { responseType: "json" });
        },
        1000,
        3
      );
      if (fileResponse && fileResponse.data) {
        return fileResponse.data;
      }
    } catch (e) {
      throw new AccessGithubError(url, "OfficeSampleProvider", e);
    }
  }
}

export const officeSampleProvider = new OfficeSampleProvider();
