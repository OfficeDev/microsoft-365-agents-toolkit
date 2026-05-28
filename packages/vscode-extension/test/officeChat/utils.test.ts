import * as chai from "chai";
import * as sinon from "sinon";
import chaiPromised from "chai-as-promised";
import * as vscode from "vscode";
import * as utils from "../../src/officeChat/utils";
import * as chatUtils from "../../src/chat/utils";
import * as dynamicPrompt from "../../src/officeChat/dynamicPrompt";
import { CancellationToken } from "../mocks/vsc";
import { AxiosResponse } from "axios";
import { Spec } from "../../src/officeChat/common/skills/spec";
import { OfficeChatTelemetryData } from "../../src/officeChat/telemetry";
import * as requestUtils from "@microsoft/teamsfx-core/build/common/requestUtils";

chai.use(chaiPromised);

describe("File: officeChat/utils.ts", () => {
  const sandbox = sinon.createSandbox();

  describe("Method: purifyUserMessage", () => {
    let officeChatTelemetryDataMock: any;
    beforeEach(() => {
      officeChatTelemetryDataMock = sandbox.createStubInstance(OfficeChatTelemetryData);
      officeChatTelemetryDataMock.chatMessages = [];
      officeChatTelemetryDataMock.responseChatMessages = [];
    });

    afterEach(() => {
      sandbox.restore();
    });

    it("purify user message successfully", async () => {
      const token = new CancellationToken();
      const getCopilotResponseAsStringStub = sandbox
        .stub(chatUtils, "getCopilotResponseAsString")
        .resolves("purified message");
      const result = await utils.purifyUserMessage("test", token, officeChatTelemetryDataMock);
      chai.assert.isTrue(getCopilotResponseAsStringStub.calledOnce);
      chai.expect(result).equal("purified message");
    });

    it("purify user message successfully", async () => {
      const token = new CancellationToken();
      const getCopilotResponseAsStringStub = sandbox
        .stub(chatUtils, "getCopilotResponseAsString")
        .resolves("");
      const result = await utils.purifyUserMessage("test", token, officeChatTelemetryDataMock);
      chai.assert.isTrue(getCopilotResponseAsStringStub.calledOnce);
      chai.expect(result).equal("test");
    });
  });

  describe("Method: isInputHarmful", () => {
    let officeChatTelemetryDataMock: any;
    beforeEach(() => {
      sandbox.stub(dynamicPrompt, "buildDynamicPrompt").returns({
        messages: [],
        version: "0.0.1",
      });
      officeChatTelemetryDataMock = sandbox.createStubInstance(OfficeChatTelemetryData);
      officeChatTelemetryDataMock.chatMessages = [];
      officeChatTelemetryDataMock.responseChatMessages = [];
    });
    afterEach(() => {
      sandbox.restore();
    });

    it("check the input is harmful", async () => {
      sandbox.stub(chatUtils, "getCopilotResponseAsString").resolves('{"isHarmful": true}```');
      const token = new CancellationToken();
      const result = await utils.isInputHarmful(
        { prompt: "test" } as unknown as vscode.ChatRequest,
        token,
        officeChatTelemetryDataMock
      );
      chai.assert.isTrue(result);
    });

    it("check the input is harmless", async () => {
      sandbox.stub(chatUtils, "getCopilotResponseAsString").resolves('{"isHarmful": false}');
      const token = new CancellationToken();
      const result = await utils.isInputHarmful(
        { prompt: "test" } as unknown as vscode.ChatRequest,
        token,
        officeChatTelemetryDataMock
      );
      chai.assert.isFalse(result);
    });

    it("get empty response", async () => {
      sandbox.stub(chatUtils, "getCopilotResponseAsString").resolves(undefined);
      const token = new CancellationToken();
      const officeChatTelemetryDataMock = sandbox.createStubInstance(OfficeChatTelemetryData);
      officeChatTelemetryDataMock.chatMessages = [];
      officeChatTelemetryDataMock.responseChatMessages = [];
      try {
        await utils.isInputHarmful(
          { prompt: "test" } as unknown as vscode.ChatRequest,
          token,
          officeChatTelemetryDataMock
        );
        chai.assert.fail("Should not reach here.");
      } catch (error) {
        chai.expect((error as Error).message).equal("Got empty response");
      }
    });

    it("isHarmful is not boolean", async () => {
      sandbox.stub(chatUtils, "getCopilotResponseAsString").resolves('{"isHarmful": "test"}');
      const token = new CancellationToken();
      const officeChatTelemetryDataMock = sandbox.createStubInstance(OfficeChatTelemetryData);
      officeChatTelemetryDataMock.chatMessages = [];
      officeChatTelemetryDataMock.responseChatMessages = [];
      try {
        await utils.isInputHarmful(
          { prompt: "test" } as unknown as vscode.ChatRequest,
          token,
          officeChatTelemetryDataMock
        );
        chai.assert.fail("Should not reach here.");
      } catch (error) {
        chai
          .expect((error as Error).message)
          .equal("Failed to parse response: isHarmful is not a boolean.");
      }
    });
  });

  describe("Method: isOutputHarmful", () => {
    beforeEach(() => {
      sandbox.stub(dynamicPrompt, "buildDynamicPrompt").returns({
        messages: [],
        version: "0.0.1",
      });
    });
    afterEach(() => {
      sandbox.restore();
    });

    it("output is harmful", async () => {
      sandbox.stub(chatUtils, "getCopilotResponseAsString").resolves("");
      const token = new CancellationToken();
      const spec = new Spec("Some user input");
      const result = await utils.isOutputHarmful("test", token, spec);
      chai.assert.isTrue(result);
    });

    it("output is harmless", async () => {
      sandbox.stub(chatUtils, "getCopilotResponseAsString").resolves("0");
      const token = new CancellationToken();
      const spec = new Spec("Some user input");
      const result = await utils.isOutputHarmful("test", token, spec);
      chai.assert.isFalse(result);
    });
  });

  describe("Method: getOfficeSample", () => {
    const date = new Date("2024-03-15T00:00:00.000Z");
    const fakedOfficeSampleConfig = {
      filterOptions: {
        capabilities: ["Excel"],
        languages: ["TS"],
        technologies: ["Office Add-in"],
      },
      samples: [
        {
          configuration: "Ready for debug",
          downloadUrlInfo: {
            owner: "OfficeDev",
            repository: "Office-Samples",
            ref: "agent",
            dir: "Excel-Add-in-ShapeAPI-Dashboard",
          },
          id: "Excel-Add-in-ShapeAPI-Dashboard",
          title: "Using shape API to work as a dashboard",
          shortDescription: "Using Shape related APIs to insert and format to work as a dashboard.",
          fullDescription:
            "The sample add-in demonstrates Excel add-in capablities to help users using shape API to work as a dashboard.",
          tags: ["TEST tag"],
          time: "5min to run",
          thumbnailPath: "assets/thumbnail.png",
          suggested: false,
          gifUrl:
            "https://raw.githubusercontent.com/OfficeDev/Office-Samples/agent/Excel-Add-in-ShapeAPI-Dashboard/assets/sampleDemo.gif",
          gifPath: "assets/sampleDemo.gif",
          onboardDate: date,
          shortId: "Shape API dashboard",
          types: ["Excel"],
        },
      ],
    };
    beforeEach(() => {
      sandbox
        .stub(requestUtils, "sendRequestWithTimeout")
        .resolves({ data: fakedOfficeSampleConfig } as AxiosResponse);
    });
    afterEach(() => {
      sandbox.restore();
    });

    it("get office sample info", async () => {
      const result = await utils.getOfficeSample("Excel-Add-in-ShapeAPI-Dashboard");
      const sample = fakedOfficeSampleConfig.samples[0];
      chai.expect(result).deep.equal(sample);
    });

    it("sample not found", async () => {
      try {
        await utils.getOfficeSample("test");
        chai.assert.fail("Should not reach here.");
      } catch (error) {
        chai.expect((error as Error).message).equal("Sample not found");
      }
    });
  });
});
