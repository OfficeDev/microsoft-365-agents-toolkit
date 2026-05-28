import { sampleProvider } from "@microsoft/teamsfx-core";
import * as chai from "chai";
import chaiPromised from "chai-as-promised";
import * as sinon from "sinon";
import * as vscode from "vscode";
import * as utils from "../../src/chat/utils";
import { CancellationToken } from "../mocks/vsc";
import * as vscodeMocks from "../mocks/vsc";
import { Tokenizer } from "../../src/chat/tokenizer";
import {
  BaseTokensPerCompletion,
  BaseTokensPerMessage,
  BaseTokensPerName,
} from "../../src/chat/consts";

chai.use(chaiPromised);

describe("chat utils", () => {
  afterEach(() => {
    // Restore the default sandbox here
    sinon.restore();
  });

  describe("verbatimCopilotInteraction()", () => {
    const sandbox = sinon.createSandbox();

    afterEach(async () => {
      sandbox.restore();
    });

    it("outputs result from LLM", async () => {
      const asyncIterator = (async function* () {
        yield "result";
      })();
      const token = new CancellationToken();
      const chatModel: vscode.LanguageModelChat = {
        sendRequest: sandbox.stub().resolves({
          text: asyncIterator,
        }),
        id: "",
        vendor: "",
        name: "",
        family: "gpt-3.5-turbo",
        version: "",
        maxInputTokens: 0,
        countTokens: sandbox.stub(),
      };
      sandbox.stub(vscode.lm, "selectChatModels").resolves([chatModel]);
      const response = {
        markdown: sandbox.stub(),
      };

      await utils.verbatimCopilotInteraction(
        "copilot-gpt-3.5-turbo",
        [],
        response as unknown as vscode.ChatResponseStream,
        token
      );
      chai.assert.isTrue(response.markdown.calledOnceWith("result"));

      await chai.assert.isRejected(
        utils.verbatimCopilotInteraction(
          "copilot-gpt-4",
          [],
          response as unknown as vscode.ChatResponseStream,
          token
        ),
        "No chat models available for the specified family"
      );
    });
  });

  describe("getCopilotResponseAsString()", () => {
    const sandbox = sinon.createSandbox();

    afterEach(async () => {
      sandbox.restore();
    });

    it("returns result as string from LLM", async () => {
      const asyncIterator = (async function* () {
        yield "result";
      })();
      const token = new CancellationToken();
      const chatModel: vscode.LanguageModelChat = {
        sendRequest: sandbox.stub().resolves({
          text: asyncIterator,
        }),
        id: "",
        vendor: "",
        name: "",
        family: "gpt-3.5-turbo",
        version: "",
        maxInputTokens: 0,
        countTokens: sandbox.stub(),
      };
      sandbox.stub(vscode.lm, "selectChatModels").resolves([chatModel]);
      const response = {
        markdown: sandbox.stub(),
      };

      const result = await utils.getCopilotResponseAsString("copilot-gpt-3.5-turbo", [], token);
      chai.assert.equal(result, "result");

      await chai.assert.isRejected(
        utils.getCopilotResponseAsString("copilot-gpt-4", [], token),
        "No chat models available for the specified family"
      );
    });
  });

  describe("getSampleDownloadUrlInfo()", () => {
    const sandbox = sinon.createSandbox();

    afterEach(async () => {
      sandbox.restore();
    });

    it("returns download Url", async () => {
      const testDownloadUrlInfo = {
        owner: "test",
        repository: "test",
        ref: "test",
        dir: "test",
      };
      sandbox.stub(sampleProvider, "SampleCollection").get(() => {
        return Promise.resolve({
          samples: [
            {
              id: "sampleId",
              downloadUrlInfo: testDownloadUrlInfo,
            },
          ],
        });
      });
      const result = await utils.getSampleDownloadUrlInfo("sampleId");
      chai.assert.equal(result, testDownloadUrlInfo);
    });

    it("throws error if not found", async () => {
      sandbox.stub(sampleProvider, "SampleCollection").get(() => {
        return Promise.resolve({
          samples: [
            {
              id: "sampleId2",
              downloadUrlInfo: undefined,
            },
          ],
        });
      });
      chai
        .expect(utils.getSampleDownloadUrlInfo("sampleId"))
        .to.be.rejectedWith("Sample not found");
    });
  });

  describe("countMessageTokens()", () => {
    const sandbox = sinon.createSandbox();

    beforeEach(() => {
      sandbox.stub(Tokenizer.getInstance(), "tokenLength").callsFake((content): number => {
        return content.length;
      });
    });

    afterEach(() => {
      sandbox.restore();
    });

    it("count empty message", () => {
      const message = new vscodeMocks.chat.LanguageModelChatMessage(
        vscodeMocks.chat.LanguageModelChatMessageRole.User,
        ""
      );
      const result = utils.countMessageTokens(message);
      chai.assert.equal(result, BaseTokensPerMessage);
    });

    it("count message without name", () => {
      const message = new vscodeMocks.chat.LanguageModelChatMessage(
        vscodeMocks.chat.LanguageModelChatMessageRole.User,
        "testContent1"
      );
      const result = utils.countMessageTokens(message);
      chai.assert.equal(result, BaseTokensPerMessage + "testContent1".length);
    });

    it("count message with name", () => {
      const message = new vscodeMocks.chat.LanguageModelChatMessage(
        vscodeMocks.chat.LanguageModelChatMessageRole.User,
        "testContent2",
        "testName2"
      );
      const result = utils.countMessageTokens(message);
      chai.assert.equal(
        result,
        BaseTokensPerMessage + "testContent2".length + "testName2".length + BaseTokensPerName
      );
    });
  });

  describe("countMessagesTokens()", () => {
    const sandbox = sinon.createSandbox();

    beforeEach(() => {
      sandbox.stub(Tokenizer.getInstance(), "tokenLength").callsFake((content): number => {
        return content.length;
      });
    });

    afterEach(() => {
      sandbox.restore();
    });

    it("count empty messages", () => {
      const messages = [] as vscodeMocks.chat.LanguageModelChatMessage[];
      const result = utils.countMessagesTokens(messages);
      chai.assert.equal(result, BaseTokensPerCompletion);
    });

    it("count messages", () => {
      const messages = [
        new vscodeMocks.chat.LanguageModelChatMessage(
          vscodeMocks.chat.LanguageModelChatMessageRole.User,
          "testContent1"
        ),
        new vscodeMocks.chat.LanguageModelChatMessage(
          vscodeMocks.chat.LanguageModelChatMessageRole.User,
          "testContent2",
          "testName2"
        ),
      ];
      const result = utils.countMessagesTokens(messages);
      chai.assert.equal(
        result,
        BaseTokensPerMessage +
          "testContent1".length +
          BaseTokensPerMessage +
          "testContent2".length +
          "testName2".length +
          BaseTokensPerName +
          BaseTokensPerCompletion
      );
    });
  });
});
