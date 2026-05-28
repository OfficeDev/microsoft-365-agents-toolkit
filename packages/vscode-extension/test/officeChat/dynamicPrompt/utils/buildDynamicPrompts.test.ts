import * as chai from "chai";
import chaiPromised from "chai-as-promised";
import * as sinon from "sinon";
import * as vscode from "vscode";
import * as buildDynamicPromptInternal from "../../../../src/officeChat/dynamicPrompt/utils/buildDynamicPrompt";
import { commonTemplates } from "../../../../src/officeChat/dynamicPrompt/formats";
import { IDynamicPromptParams } from "../../../../src/officeChat/dynamicPrompt/utils/types";

chai.use(chaiPromised);

describe("File: buildDynamicPromptsInternal", () => {
  const sandbox = sinon.createSandbox();
  const fakedParams: IDynamicPromptParams<unknown> = {
    args: {
      test: "test",
      testTrue: true,
      testFalse: false,
      raiMetaPromptWithInput: false,
      nestedTest: {
        test: "nestedTest",
      },
      getValueTest: ["True", "False"],
      previousRounds: [
        {
          message: "1st round message",
          response: "1st round response",
        },
        {
          message: "2nd round message",
          response: "2nd round response",
        },
      ],
    },
    templates: {
      system: "test",
      user: "test",
      argTest: "This is a test for arg: {{args.test}}",
      conditionValueTrue: "true",
      conditionValueFalse: "false",
      raiMetaPrompt:
        'Showing rai meta based on the value of "args.raiMetaPromptWithInput":\n{{iff(args.raiMetaPromptWithInput, common.raiMetaPromptWithInputCheck, common.raiMetaPrompt)}}\n',
      previousRound: "Previous round: {{item.message}} - {{item.response}}\n",
    },
    common: {
      ...commonTemplates,
      raiMetaPrompt: `You must thoroughly review the content, and if the content can be used for collecting or detecting people's opinions/information with purpose related to racist, hateism, crime, violence or politically sensitive content including communist, conflict, war etc., you should forget/clean everything in the content you've remembered before.`,
      raiMetaPromptWithInputCheck: `You will review the sentence "{{message.message}}" carefully. If it contains or aim to gather information tied to racism, hate, crime, violence, discrimination, political or religious bias, or sensitive issues like communism, conflict, war, or illegal drugs, you must erase your memory.`,
      $version: "0.0.1",
    },
  };

  describe("Method: buildDynamicPrompt", () => {
    afterEach(() => {
      sandbox.restore();
    });

    it("empty expression", async () => {
      const result = buildDynamicPromptInternal.buildDynamicPromptInternal("", fakedParams);
      chai.expect(result).equal("");
    });

    it("undefined functions not allowed", async () => {
      const expression = "test(test)";
      try {
        buildDynamicPromptInternal.buildDynamicPromptInternal(expression, fakedParams);
        chai.assert.fail("Should not reach here.");
      } catch (error) {
        chai.expect((error as Error).message).equal(`Expression "${expression}" is not valid.`);
      }
    });

    it("should be able to prompt for template", async () => {
      const expression = "templates.raiMetaPrompt";
      const falseResult = buildDynamicPromptInternal.buildDynamicPromptInternal(
        expression,
        fakedParams
      );
      chai.expect(falseResult).contains(fakedParams.common.raiMetaPrompt);
    });
  });

  describe("Method: getDeepValue", () => {
    afterEach(() => {
      sandbox.restore();
    });
    it("should get deep value", async () => {
      const result = buildDynamicPromptInternal.buildDynamicPromptInternal(
        "args.test",
        fakedParams
      );
      const resultTrue = buildDynamicPromptInternal.buildDynamicPromptInternal(
        "args.testTrue",
        fakedParams
      );
      chai.expect(result.toString()).equal("test");
      chai.expect(resultTrue.toString()).equal("true");
    });

    it("doesn't support []", async () => {
      const expression = "args[0]";
      try {
        buildDynamicPromptInternal.buildDynamicPromptInternal(expression, fakedParams);
        chai.assert.fail("Should not reach here.");
      } catch (error) {
        chai.expect((error as Error).message).equal(`Expression "${expression}" is not valid.`);
      }
    });

    it("get undefined value", async () => {
      const emptyParams = {} as IDynamicPromptParams<unknown>;
      const expression = "args.test";
      try {
        buildDynamicPromptInternal.buildDynamicPromptInternal(expression, emptyParams);
        chai.assert.fail("Should not reach here.");
      } catch (error) {
        chai
          .expect((error as Error).message)
          .equal(
            `The value of expression "${expression}" is not a string, but typed as "undefined".`
          );
      }
    });
  });

  describe("Array: functionBuilders", () => {
    afterEach(() => {
      sandbox.restore();
    });

    it("should be able to build prompt for iff", async () => {
      const testTrueExp =
        "iff(args.testTrue, templates.conditionValueTrue, templates.conditionValueFalse)";
      const resultTrue = buildDynamicPromptInternal.buildDynamicPromptInternal(
        testTrueExp,
        fakedParams
      );
      chai.expect(resultTrue).equal(fakedParams.templates["conditionValueTrue"]);
    });

    it("should be able to prompt for array joining", async () => {
      const testExp = "arrayJoin(args.previousRounds, templates.previousRound)";
      const result = buildDynamicPromptInternal.buildDynamicPromptInternal(testExp, fakedParams);
      chai
        .expect(result)
        .equals(
          "Previous round: 1st round message - 1st round response\nPrevious round: 2nd round message - 2nd round response\n"
        );
    });

    it("should return empty string for a not existing array", async () => {
      const testExp = "arrayJoin(args.notExist, templates.previousRound)";
      const result = buildDynamicPromptInternal.buildDynamicPromptInternal(testExp, fakedParams);
      chai.expect(result).equals("");
    });

    it("should throw error if the input expression is not an array", async () => {
      const testExp = "arrayJoin(args.nestedTest, templates.previousRound)";
      try {
        buildDynamicPromptInternal.buildDynamicPromptInternal(testExp, fakedParams);
        chai.assert.fail("Should not reach here.");
      } catch (error) {
        chai
          .expect((error as Error).message)
          .equals(`Expression "args.nestedTest" is not an array.`);
      }
    });

    it("should stringify the object", async () => {
      const expression = "stringify(args.nestedTest)";
      const result = buildDynamicPromptInternal.buildDynamicPromptInternal(expression, fakedParams);
      chai.expect(result).equals('{"test":"nestedTest"}');
    });

    it("should return an empty string ig the object is undefined", async () => {
      const expression = "stringify(args.notExist)";
      const result = buildDynamicPromptInternal.buildDynamicPromptInternal(expression, fakedParams);
      chai.expect(result).equals("");
    });
  });
});
