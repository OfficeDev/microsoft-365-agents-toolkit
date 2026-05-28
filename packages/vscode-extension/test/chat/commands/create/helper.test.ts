import { sampleProvider } from "@microsoft/teamsfx-core";
import * as generatorUtils from "@microsoft/teamsfx-core/build/component/generator/utils";
import axios from "axios";
import * as chai from "chai";
import chaiPromised from "chai-as-promised";
import fs from "fs-extra";
import path from "path";
import * as sinon from "sinon";
import tmp from "tmp";
import * as vscode from "vscode";
import * as helper from "../../../../src/chat/commands/create/helper";
import { ProjectMetadata } from "../../../../src/chat/commands/create/types";
import * as telemetry from "../../../../src/chat/telemetry";
import * as util from "../../../../src/chat/utils";
import { ExtTelemetry } from "../../../../src/telemetry/extTelemetry";
import { CancellationToken } from "../../../mocks/vsc";

chai.use(chaiPromised);

describe("chat create helper", () => {
  afterEach(() => {
    sinon.restore();
  });

  describe("matchProject()", () => {
    const sandbox = sinon.createSandbox();

    afterEach(async () => {
      sandbox.restore();
    });

    it("has matched sample project", async () => {
      const chatTelemetryDataMock = sandbox.createStubInstance(telemetry.ChatTelemetryData);
      sandbox.stub(chatTelemetryDataMock, "properties").get(function getterFn() {
        return undefined;
      });
      sandbox.stub(chatTelemetryDataMock, "measurements").get(function getterFn() {
        return undefined;
      });
      sandbox.stub(sampleProvider, "SampleCollection").get(function getterFn() {
        return {
          samples: [
            {
              id: "test1",
              title: "test1",
              fullDescription: "test1",
            },
          ],
        };
      });
      chatTelemetryDataMock.chatMessages = [];
      sandbox
        .stub(telemetry.ChatTelemetryData, "createByParticipant")
        .returns(chatTelemetryDataMock);
      const sendTelemetryEventStub = sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
      sandbox
        .stub(util, "getCopilotResponseAsString")
        .onFirstCall()
        .resolves('{"app":[{"id": "test1", "score": 1.0}]}')
        .onSecondCall()
        .resolves('{"app":[{"id": "test2", "score": 0.5}]}');

      const token = new CancellationToken();
      const result = await helper.matchProject(
        { prompt: "test" } as vscode.ChatRequest,
        token,
        chatTelemetryDataMock
      );
      chai.assert.strictEqual(result.length, 1);
      chai.assert.strictEqual(result[0].id, "test1");
    });

    it("has matched template project", async () => {
      const chatTelemetryDataMock = sandbox.createStubInstance(telemetry.ChatTelemetryData);
      sandbox.stub(chatTelemetryDataMock, "properties").get(function getterFn() {
        return undefined;
      });
      sandbox.stub(chatTelemetryDataMock, "measurements").get(function getterFn() {
        return undefined;
      });
      sandbox.stub(sampleProvider, "SampleCollection").get(function getterFn() {
        return {
          samples: [
            {
              id: "test1",
              title: "test1",
              fullDescription: "test1",
            },
          ],
        };
      });
      chatTelemetryDataMock.chatMessages = [];
      sandbox
        .stub(telemetry.ChatTelemetryData, "createByParticipant")
        .returns(chatTelemetryDataMock);
      const sendTelemetryEventStub = sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
      sandbox
        .stub(util, "getCopilotResponseAsString")
        .onFirstCall()
        .resolves('{"app":[{"id": "bot", "score": 1.0}]}')
        .onSecondCall()
        .resolves('{"app":[{"id": "test2", "score": 0.5}]}');

      const token = new CancellationToken();
      const result = await helper.matchProject(
        { prompt: "test template" } as vscode.ChatRequest,
        token,
        chatTelemetryDataMock
      );
      chai.assert.strictEqual(result.length, 1);
      chai.assert.strictEqual(result[0].id, "bot");
    });
  });

  describe("showFileTree()", () => {
    const sandbox = sinon.createSandbox();
    afterEach(async () => {
      sandbox.restore();
    });

    it("calls filetree API", async () => {
      sandbox.stub(util, "getSampleDownloadUrlInfo").resolves({
        owner: "test",
        repository: "testRepo",
        ref: "testRef",
        dir: "testDir",
      });
      sandbox.stub(generatorUtils, "getSampleFileInfo").resolves({
        samplePaths: ["test"],
        fileUrlPrefix: "https://test.com/",
      });
      sandbox.stub(tmp, "dirSync").returns({
        name: "tempDir",
      } as unknown as tmp.DirResult);
      sandbox.stub(axios, "get").callsFake(async (url: string, config) => {
        if (url === "https://test.com/test") {
          return { data: "testData", status: 200 };
        } else {
          throw new Error("Invalid URL");
        }
      });
      sandbox.stub(fs, "ensureFile");
      sandbox.stub(fs, "writeFile");

      const projectMetadata = {
        id: "test1",
        type: "sample",
        platform: "Teams",
        name: "test1",
        description: "test1",
      } as ProjectMetadata;
      const response = {
        markdown: sandbox.stub(),
        filetree: sandbox.stub(),
      };
      const result = await helper.showFileTree(
        projectMetadata,
        response as unknown as vscode.ChatResponseStream
      );
      chai.assert.isTrue(response.filetree.calledOnce);
      chai.assert.strictEqual(result, path.join("tempDir", "testDir"));
    });
  });
});
