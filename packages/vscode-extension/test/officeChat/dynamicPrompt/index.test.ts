import * as chai from "chai";
import chaiPromised from "chai-as-promised";
import * as sinon from "sinon";
import * as vscode from "vscode";
import * as buildDynamicPrompt from "../../../src/officeChat/dynamicPrompt";
import * as buildDynamicPromptInternal from "../../../src/officeChat/dynamicPrompt/utils/buildDynamicPrompt";
import { IDynamicPromptFormat } from "../../../src/officeChat/dynamicPrompt/utils/types";

chai.use(chaiPromised);

describe("File: dynamicPrompt/index", () => {
  const sandbox = sinon.createSandbox();

  describe("Methos: buildDynamicPrompt", () => {
    const fakedFormat: IDynamicPromptFormat<string> = {
      templates: {
        system: "test",
        user: "test",
      },
      messages: [
        {
          role: "system",
          entryTemplate: "system",
        },
        {
          role: "user",
          entryTemplate: "user",
        },
      ],
      version: "0.1",
    };
    const prompt = "test";
    afterEach(() => {
      sandbox.restore();
    });
    it("build dynamic prompts successfully", async () => {
      sandbox.stub(buildDynamicPromptInternal, "buildDynamicPromptInternal").returns("test");
      const result = buildDynamicPrompt.buildDynamicPrompt(fakedFormat, prompt);
      chai
        .expect(result.messages[0])
        .deep.equal(
          new vscode.LanguageModelChatMessage(vscode.LanguageModelChatMessageRole.User, "test")
        );
    });

    it("throw exceptions", async () => {
      sandbox
        .stub(buildDynamicPromptInternal, "buildDynamicPromptInternal")
        .throws(new Error("test error"));
      try {
        buildDynamicPrompt.buildDynamicPrompt(fakedFormat, prompt);
        chai.assert.fail("should not reach here.");
      } catch (error) {
        chai.expect((error as Error).message).equal("test error");
      }
    });

    it("create assistant message", async () => {
      const fakedAssistantFormat: IDynamicPromptFormat<string> = {
        templates: {
          assistant: "test",
        },
        messages: [
          {
            role: "assistant",
            entryTemplate: "assistant",
          },
        ],
        version: "0.1",
      };
      sandbox.stub(buildDynamicPromptInternal, "buildDynamicPromptInternal").returns("test");
      try {
        const result = buildDynamicPrompt.buildDynamicPrompt(fakedAssistantFormat, prompt);
        chai
          .expect(result.messages[0])
          .deep.equal(
            new vscode.LanguageModelChatMessage(
              vscode.LanguageModelChatMessageRole.Assistant,
              "test"
            )
          );
      } catch (error) {
        chai.expect((error as Error).name).equal("TypeError");
      }
    });
  });
});
