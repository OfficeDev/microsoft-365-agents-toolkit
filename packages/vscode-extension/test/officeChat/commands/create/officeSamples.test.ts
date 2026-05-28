import * as chai from "chai";
import * as sinon from "sinon";
import axios from "axios";
import { AccessGithubError } from "@microsoft/teamsfx-core";
import { officeSampleProvider } from "../../../../src/officeChat/commands/create/officeSamples";
import { err } from "@microsoft/teamsfx-api";

describe("File: officeSamples", () => {
  const sandbox = sinon.createSandbox();
  const fakedOfficeSampleConfig = {
    filterOptions: {
      capabilities: ["Excel"],
      languages: ["TS"],
      technologies: ["Office Add-in"],
    },
    samples: [
      {
        id: "Excel-Add-in-ShapeAPI-Dashboard",
        title: "Using shape API to work as a dashboard",
        shortDescription: "Using Shape related APIs to insert and format to work as a dashboard.",
        fullDescription:
          "The sample add-in demonstrates Excel add-in capablities to help users using shape API to work as a dashboard.",
        tags: ["TS", "Shape", "Excel", "Office Add-in"],
        time: "5min to run",
        configuration: "Ready for debug",
        thumbnailPath: "",
        suggested: false,
      },
    ],
  };
  const fakedOfficeSampleConfigWithGif = {
    filterOptions: {
      capabilities: ["Excel"],
      languages: ["TS"],
      technologies: ["Office Add-in"],
    },
    samples: [
      {
        id: "Excel-Add-in-ShapeAPI-Dashboard",
        title: "Using shape API to work as a dashboard",
        shortDescription: "Using Shape related APIs to insert and format to work as a dashboard.",
        fullDescription:
          "The sample add-in demonstrates Excel add-in capablities to help users using shape API to work as a dashboard.",
        tags: ["TS", "Shape", "Excel", "Office Add-in"],
        time: "5min to run",
        configuration: "Ready for debug",
        thumbnailPath: "",
        gifPath: "assets/sampleDemo.gif",
        suggested: false,
      },
    ],
  };

  afterEach(() => {
    sandbox.restore();
    officeSampleProvider["officeSampleCollection"] = undefined;
  });

  it("download office sample config", async () => {
    sandbox.stub(axios, "get").callsFake(async (url: string, config) => {
      if (
        url ===
        "https://raw.githubusercontent.com/OfficeDev/Office-Samples/agent/.config/samples-config-v1.json"
      ) {
        return { data: fakedOfficeSampleConfig, status: 200 };
      } else {
        throw err(undefined);
      }
    });
    const samples = (await officeSampleProvider.OfficeSampleCollection).samples;
    chai.expect(samples[0].downloadUrlInfo).deep.equal({
      owner: "OfficeDev",
      repository: "Office-Samples",
      ref: "agent",
      dir: "Excel-Add-in-ShapeAPI-Dashboard",
    });
    chai.expect(samples[0].gifUrl).equal(undefined);
  });

  it("download office sample config with gif link", async () => {
    sandbox.stub(axios, "get").callsFake(async (url: string, config) => {
      if (
        url ===
        "https://raw.githubusercontent.com/OfficeDev/Office-Samples/agent/.config/samples-config-v1.json"
      ) {
        return { data: fakedOfficeSampleConfigWithGif, status: 200 };
      } else {
        throw err(undefined);
      }
    });
    const samples = (await officeSampleProvider.OfficeSampleCollection).samples;
    chai.expect(samples[0].downloadUrlInfo).deep.equal({
      owner: "OfficeDev",
      repository: "Office-Samples",
      ref: "agent",
      dir: "Excel-Add-in-ShapeAPI-Dashboard",
    });
    chai
      .expect(samples[0].gifUrl)
      .equal(
        `https://raw.githubusercontent.com/OfficeDev/Office-Samples/agent/Excel-Add-in-ShapeAPI-Dashboard/assets/sampleDemo.gif`
      );
  });

  it("online sample config returns undefined when failed to fetch", async () => {
    sandbox.stub(axios, "get").callsFake(async (url: string, config) => {
      if (
        url !==
        "https://raw.githubusercontent.com/OfficeDev/Office-Samples/agent/.config/samples-config-v1.json"
      ) {
        throw new Error("test error");
      }
    });

    try {
      await officeSampleProvider.OfficeSampleCollection;
      chai.assert.fail("should not reach here");
    } catch (e) {
      chai.assert.isTrue(e instanceof AccessGithubError);
    }
  });

  it("SampleCollection already exists", async () => {
    const providerInstance = officeSampleProvider;
    providerInstance["officeSampleCollection"] = {
      samples: [],
      fileterOptions: {
        capabilities: ["Excel"],
        languages: ["TS"],
        technologies: ["Office Add-in"],
      },
    };
    const fileterOptions = (await officeSampleProvider.OfficeSampleCollection).fileterOptions;
    chai.expect(fileterOptions.capabilities[0]).equal("Excel");
  });
});
