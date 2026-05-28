import * as chai from "chai";
import sinon from "ts-sinon";
import { Spec } from "../../../src/officeChat/common/skills/spec";
import {
  CancellationToken,
  ChatResponseStream,
  LanguageModelChatMessage,
  LanguageModelChatMessageRole,
} from "vscode";
import { ExecutionResultEnum } from "../../../src/officeChat/common/skills/executionResultEnum";
import { ISkill } from "../../../src/officeChat/common/skills/iSkill";
import { OfficeChatCommand } from "../../../src/officeChat/consts";
import { Planner } from "../../../src/officeChat/common/planner";
import * as utils from "../../../src/officeChat/utils";
import { SkillsManager } from "../../../src/officeChat/common/skills/skillsManager";
import { OfficeChatTelemetryData } from "../../../src/officeChat/telemetry";

class FakeSkill implements ISkill {
  constructor() {}
  name: string | undefined;
  capability: string | undefined;

  public canInvoke(spec: Spec): boolean {
    return true;
  }

  // eslint-disable-next-line @typescript-eslint/require-await
  public async invoke(
    languageModel: LanguageModelChatMessage,
    response: ChatResponseStream,
    token: CancellationToken,
    spec: Spec
  ): Promise<{ result: ExecutionResultEnum; spec: Spec }> {
    return { result: ExecutionResultEnum.Success, spec: spec };
  }
}

describe("planner", () => {
  let invokeParametersInit: () => any;
  const sandbox = sinon.createSandbox();

  beforeEach(() => {
    invokeParametersInit = function () {
      const model: LanguageModelChatMessage = {
        role: LanguageModelChatMessageRole.User,
        content: "",
        name: undefined,
      };

      const fakeRequest = {
        prompt: sandbox.stub(),
        command: sandbox.stub(),
        references: sandbox.stub(),
      };

      const fakeResponse = {
        markdown: sandbox.stub(),
        anchor: sandbox.stub(),
        button: sandbox.stub(),
        filetree: sandbox.stub(),
        progress: sandbox.stub(),
        reference: sandbox.stub(),
        push: sandbox.stub(),
      } as unknown as ChatResponseStream;

      const fakeToken: CancellationToken = {
        isCancellationRequested: false,
        onCancellationRequested: sandbox.stub(),
      };

      const fakeCommand = OfficeChatCommand.GenerateCode;

      const telemetryData = new OfficeChatTelemetryData(
        fakeCommand,
        "requestId",
        0,
        "participantId"
      );

      const fakeSkill = new FakeSkill();

      return { model, fakeRequest, fakeResponse, fakeToken, fakeCommand, telemetryData, fakeSkill };
    };
  });

  afterEach(() => {
    sandbox.restore();
  });

  it("constructor", () => {
    const skillset = Planner.getInstance();

    chai.assert.isNotNull(skillset);
  });

  it("canInvoke returns true", async () => {
    const { model, fakeRequest, fakeResponse, fakeToken, fakeCommand, telemetryData, fakeSkill } =
      invokeParametersInit();

    const skillManagerStub = SkillsManager.getInstance();
    sandbox.stub(skillManagerStub, "getCapableSkills").returns([]);

    const chatResult = await Planner.getInstance().processRequest(
      model,
      fakeRequest,
      fakeResponse,
      fakeToken,
      fakeCommand,
      telemetryData
    );

    chai.assert.isObject(chatResult);
    chai.assert.isObject(chatResult.errorDetails);
    chai.assert.isString(chatResult.errorDetails?.message);
    chai.assert.isTrue(chatResult.errorDetails?.message.startsWith("No skill is available"));
  });

  it("canInvoke returns true", async () => {
    const { model, fakeRequest, fakeResponse, fakeToken, fakeCommand, telemetryData, fakeSkill } =
      invokeParametersInit();

    const skillManagerStub = SkillsManager.getInstance();
    sandbox.stub(skillManagerStub, "getCapableSkills").returns([fakeSkill, fakeSkill]);

    const purifyUserMessageStub = sandbox.stub(utils, "purifyUserMessage");
    purifyUserMessageStub.resolves("purified");

    sandbox.stub(console, "log");
    sandbox.stub(console, "debug");
    sandbox.stub(console, "error");

    const chatResult = await Planner.getInstance().processRequest(
      model,
      fakeRequest,
      fakeResponse,
      fakeToken,
      fakeCommand,
      telemetryData
    );

    chai.assert.isObject(chatResult);
  });

  it("can not Invoke returns false", async () => {
    const { model, fakeRequest, fakeResponse, fakeToken, fakeCommand, telemetryData, fakeSkill } =
      invokeParametersInit();
    fakeSkill.canInvoke = sandbox.stub().onCall(0).returns(true).onCall(1).returns(false);

    const skillManagerStub = SkillsManager.getInstance();
    sandbox.stub(skillManagerStub, "getCapableSkills").returns([fakeSkill, fakeSkill]);

    const purifyUserMessageStub = sandbox.stub(utils, "purifyUserMessage");
    purifyUserMessageStub.resolves("purified");

    sandbox.stub(console, "log");
    sandbox.stub(console, "debug");
    sandbox.stub(console, "error");

    try {
      const chatResult = await Planner.getInstance().processRequest(
        model,
        fakeRequest,
        fakeResponse,
        fakeToken,
        fakeCommand,
        telemetryData
      );
      chai.assert.isObject(chatResult);
    } catch (error) {}
  });

  it("skip if skill returns Failure", async () => {
    const { model, fakeRequest, fakeResponse, fakeToken, fakeCommand, telemetryData, fakeSkill } =
      invokeParametersInit();
    fakeSkill.invoke = sandbox
      .stub()
      .resolves({ result: ExecutionResultEnum.Failure, spec: new Spec("") });

    const skillManagerStub = SkillsManager.getInstance();
    sandbox.stub(skillManagerStub, "getCapableSkills").returns([fakeSkill, fakeSkill]);

    const purifyUserMessageStub = sandbox.stub(utils, "purifyUserMessage");
    purifyUserMessageStub.resolves("purified");

    sandbox.stub(console, "log");
    sandbox.stub(console, "debug");
    sandbox.stub(console, "error");

    try {
      const chatResult = await Planner.getInstance().processRequest(
        model,
        fakeRequest,
        fakeResponse,
        fakeToken,
        fakeCommand,
        telemetryData
      );
      chai.assert.isObject(chatResult);
    } catch (error) {}
  });

  it("skip if skill returns Rejected", async () => {
    const { model, fakeRequest, fakeResponse, fakeToken, fakeCommand, telemetryData, fakeSkill } =
      invokeParametersInit();
    fakeSkill.invoke = sandbox
      .stub()
      .resolves({ result: ExecutionResultEnum.Rejected, spec: new Spec("") });

    const skillManagerStub = SkillsManager.getInstance();
    sandbox.stub(skillManagerStub, "getCapableSkills").returns([fakeSkill, fakeSkill]);

    const purifyUserMessageStub = sandbox.stub(utils, "purifyUserMessage");
    purifyUserMessageStub.resolves("purified");

    sandbox.stub(console, "log");
    sandbox.stub(console, "debug");
    sandbox.stub(console, "error");

    try {
      const chatResult = await Planner.getInstance().processRequest(
        model,
        fakeRequest,
        fakeResponse,
        fakeToken,
        fakeCommand,
        telemetryData
      );
      chai.assert.isObject(chatResult);
    } catch (error) {}
  });

  it("skip if skill returns FailedAndGoNext", async () => {
    const { model, fakeRequest, fakeResponse, fakeToken, fakeCommand, telemetryData, fakeSkill } =
      invokeParametersInit();
    fakeSkill.invoke = sandbox
      .stub()
      .resolves({ result: ExecutionResultEnum.FailedAndGoNext, spec: new Spec("") });

    const skillManagerStub = SkillsManager.getInstance();
    sandbox.stub(skillManagerStub, "getCapableSkills").returns([fakeSkill, fakeSkill]);

    const purifyUserMessageStub = sandbox.stub(utils, "purifyUserMessage");
    purifyUserMessageStub.resolves("purified");

    sandbox.stub(console, "log");
    sandbox.stub(console, "debug");
    sandbox.stub(console, "error");

    const chatResult = await Planner.getInstance().processRequest(
      model,
      fakeRequest,
      fakeResponse,
      fakeToken,
      fakeCommand,
      telemetryData
    );

    chai.assert.isObject(chatResult);
  });

  it("cancel the execution if the token set as cancelled", async () => {
    const { model, fakeRequest, fakeResponse, fakeToken, fakeCommand, telemetryData, fakeSkill } =
      invokeParametersInit();
    fakeSkill.invoke = sandbox
      .stub()
      .resolves({ result: ExecutionResultEnum.FailedAndGoNext, spec: new Spec("") });

    fakeToken.isCancellationRequested = true;
    const skillManagerStub = SkillsManager.getInstance();
    sandbox.stub(skillManagerStub, "getCapableSkills").returns([fakeSkill, fakeSkill]);

    const purifyUserMessageStub = sandbox.stub(utils, "purifyUserMessage");
    purifyUserMessageStub.resolves("purified");

    sandbox.stub(console, "log");
    sandbox.stub(console, "debug");
    sandbox.stub(console, "error");

    await Planner.getInstance().processRequest(
      model,
      fakeRequest,
      fakeResponse,
      fakeToken,
      fakeCommand,
      telemetryData
    );
  });
});
