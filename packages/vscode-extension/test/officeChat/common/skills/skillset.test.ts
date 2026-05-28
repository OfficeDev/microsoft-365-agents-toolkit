import * as chai from "chai";
import sinon from "ts-sinon";
import { Spec } from "../../../../src/officeChat/common/skills/spec";
import {
  CancellationToken,
  ChatResponseStream,
  LanguageModelChatMessage,
  LanguageModelChatMessageRole,
} from "vscode";
import { ExecutionResultEnum } from "../../../../src/officeChat/common/skills/executionResultEnum";
import { SkillSet } from "../../../../src/officeChat/common/skills/skillset";
import { ISkill } from "../../../../src/officeChat/common/skills/iSkill";
import { SampleData } from "../../../../src/officeChat/common/samples/sampleData";

describe("skillset", () => {
  let invokeParametersInit: () => any;
  const sandbox = sinon.createSandbox();

  beforeEach(() => {
    invokeParametersInit = function () {
      const spec = new Spec("some user input");
      spec.taskSummary = "some task summary";
      spec.sections = ["section1", "section2"];
      spec.inspires = ["inspire1", "inspire2"];
      spec.resources = ["resource1", "resource2"];
      spec.appendix = {
        host: "some host",
        codeSnippet: "some code",
        codeExplanation: "some explanation",
        codeTaskBreakdown: ["task1", "task2"],
        codeSample: "",
        apiDeclarationsReference: new Map<string, SampleData>(),
        isCustomFunction: false,
        telemetryData: {
          requestId: "Id",
          isHarmful: false,
          relatedSampleName: ["sample1", "sample2"],
          chatMessages: [
            new LanguageModelChatMessage(LanguageModelChatMessageRole.User, "requestMessage1"),
            new LanguageModelChatMessage(LanguageModelChatMessageRole.User, "requestMessage2"),
          ],
          responseChatMessages: [
            new LanguageModelChatMessage(LanguageModelChatMessageRole.User, "responseMessage1"),
            new LanguageModelChatMessage(LanguageModelChatMessageRole.User, "responseMessage2"),
          ],
          properties: { property1: "value1", property2: "value2" },
          measurements: { measurement1: 1, measurement2: 2 },
        },
        complexity: 0,
        shouldContinue: false,
      };

      const model: LanguageModelChatMessage = {
        role: LanguageModelChatMessageRole.User,
        content: "",
        name: undefined,
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

      return { spec, model, fakeResponse, fakeToken };
    };
  });

  afterEach(() => {
    sandbox.restore();
  });

  it("constructor", () => {
    const skillset = new SkillSet([]);

    chai.assert.isNotNull(skillset);
    chai.assert.equal(skillset.name, "skillSet");
    chai.assert.equal(skillset.capability, "A container for muultiple skills");
    chai.assert.equal(skillset.retriableTimes, 1);
  });

  it("canInvoke returns true", () => {
    const fakeSkills: ISkill[] = [
      {
        name: "Skill 1",
        capability: "Beginner",
        canInvoke: sandbox.stub(),
        invoke: sandbox.stub(),
      },
      {
        name: "Skill 2",
        capability: "Intermediate",
        canInvoke: sandbox.stub(),
        invoke: sandbox.stub(),
      },
    ];

    const fakeSpec = new Spec("some user input");

    const skillset = new SkillSet(fakeSkills);

    const result = skillset.canInvoke(fakeSpec);
    chai.assert.isTrue(result);
  });

  it("canInvoke returns false", () => {
    const fakeSpec = new Spec("some user input");

    const skillset = new SkillSet([]);
    skillset.skills = undefined;

    const result = skillset.canInvoke(fakeSpec);
    chai.assert.isFalse(result);
  });

  it("skillset Invoke success with no skills", async () => {
    const { spec, model, fakeResponse, fakeToken } = invokeParametersInit();
    const skillset = new SkillSet([]);
    skillset.skills = undefined;

    const result = await skillset.invoke(model, fakeResponse, fakeToken, spec);
    chai.expect(result.result).to.equal(ExecutionResultEnum.Success);
    chai.expect(spec).to.equal(spec);
  });

  it("skillset Invoke failure with no skill can invoke", async () => {
    const { spec, model, fakeResponse, fakeToken } = invokeParametersInit();

    const fakeSkills: ISkill[] = [
      {
        name: "Skill 1",
        capability: "Beginner",
        canInvoke: sandbox.stub().returns(false),
        invoke: sandbox.stub(),
      },
    ];

    const skillset = new SkillSet(fakeSkills, 1);

    const result = await skillset.invoke(model, fakeResponse, fakeToken, spec);
    chai.expect(result.result).to.equal(ExecutionResultEnum.Failure);
    chai.expect(spec).to.equal(spec);
  });

  it("skillset Invoke success", async () => {
    const { spec, model, fakeResponse, fakeToken } = invokeParametersInit();

    const fakeSkills: ISkill[] = [
      {
        name: "Skill 1",
        capability: "Beginner",
        canInvoke: sandbox.stub().returns(true),
        invoke: sandbox.stub().returns({ result: ExecutionResultEnum.Success, spec }),
      },
    ];

    const skillset = new SkillSet(fakeSkills, 1);

    const result = await skillset.invoke(model, fakeResponse, fakeToken, spec);
    chai.expect(result.result).to.equal(ExecutionResultEnum.Success);
    chai.expect(spec).to.equal(spec);
  });

  it("skillset Invoke rejected", async () => {
    const { spec, model, fakeResponse, fakeToken } = invokeParametersInit();

    const fakeSkills: ISkill[] = [
      {
        name: "Skill 1",
        capability: "Beginner",
        canInvoke: sandbox.stub().returns(true),
        invoke: sandbox.stub().returns({ result: ExecutionResultEnum.Rejected, spec }),
      },
    ];

    const skillset = new SkillSet(fakeSkills, 1);

    const result = await skillset.invoke(model, fakeResponse, fakeToken, spec);
    chai.expect(result.result).to.equal(ExecutionResultEnum.Rejected);
    chai.expect(spec).to.equal(spec);
  });

  it("skillset Invoke failure", async () => {
    const { spec, model, fakeResponse, fakeToken } = invokeParametersInit();

    const fakeSkills: ISkill[] = [
      {
        name: "Skill 1",
        capability: "Beginner",
        canInvoke: sandbox.stub().returns(true),
        invoke: sandbox.stub().returns({ result: ExecutionResultEnum.Failure, spec }),
      },
    ];

    const skillset = new SkillSet(fakeSkills, 1);

    const result = await skillset.invoke(model, fakeResponse, fakeToken, spec);
    chai.expect(result.result).to.equal(ExecutionResultEnum.Failure);
    chai.expect(spec).to.equal(spec);
  });

  it("skillset Invoke failed and go next", async () => {
    const { spec, model, fakeResponse, fakeToken } = invokeParametersInit();

    const fakeSkills: ISkill[] = [
      {
        name: "Skill 1",
        capability: "Beginner",
        canInvoke: sandbox.stub().returns(true),
        invoke: sandbox.stub().returns({ result: ExecutionResultEnum.Failure, spec }),
      },
      {
        name: "Skill 2",
        capability: "Beginner",
        canInvoke: sandbox.stub().returns(true),
        invoke: sandbox.stub().returns({ result: ExecutionResultEnum.FailedAndGoNext, spec }),
      },
    ];

    const skillset = new SkillSet(fakeSkills, 1);

    const result = await skillset.invoke(model, fakeResponse, fakeToken, spec);
    chai.expect(result.result).to.equal(ExecutionResultEnum.FailedAndGoNext);
    chai.expect(spec).to.equal(spec);
  });
});
