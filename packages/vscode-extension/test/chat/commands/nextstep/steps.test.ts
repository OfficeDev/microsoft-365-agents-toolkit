import * as chai from "chai";
import chaiPromised from "chai-as-promised";
import * as sinon from "sinon";
import { allSteps } from "../../../../src/chat/commands/nextstep/steps";
import * as condition from "../../../../src/chat/commands/nextstep/condition";
import { DescripitionFunc, WholeStatus } from "../../../../src/chat/commands/nextstep/types";

chai.use(chaiPromised);

const titles = {
  gettingStarted: "Getting started with Teams Toolkit",
  createOrOpenProject: "Create a new project or open an existing project",
  summarizeReadme: "Get more info about the project with README",
  previewInTestTool: "Preview in Test Tool",
  signInM365Account: "Sign in to Microsoft 365 Account",
  joinM365DeveloperProgram: "Join Microsoft 365 Developer Program",
  previewInTeams: "Preview in Microsoft Teams",
  howToExtend: "How to Extend your Teams Application Capabilities",
  ciCd: "Set up CI/CD Pipelines",
  azureAccount: "Deploy Your App using Your Azure Account",
  provision: "Provision Azure resources",
  deploy: "Deploy to Azure",
  publish: "Publish Your App",
  previewRemotely: "Preview Remotely",
};

describe("next steps", () => {
  const sandbox = sinon.createSandbox();
  const steps = allSteps();

  describe(`title: "${titles.gettingStarted}"`, () => {
    afterEach(() => {
      sandbox.restore();
    });

    it("condition: selected", () => {
      sandbox.stub(condition, "isFirstInstalled").returns(true);
      const step = steps.find((s) => s.title === titles.gettingStarted);
      chai.assert.isNotEmpty(step);
      chai.assert.isTrue(step?.condition({} as WholeStatus));
    });

    it("condition: not selected", () => {
      sandbox.stub(condition, "isFirstInstalled").returns(false);
      const step = steps.find((s) => s.title === titles.gettingStarted);
      chai.assert.isFalse(step?.condition({} as WholeStatus));
    });
  });

  describe(`title: "${titles.createOrOpenProject}"`, () => {
    afterEach(() => {
      sandbox.restore();
    });

    it("condition: selected", () => {
      sandbox.stub(condition, "isProjectOpened").returns(false);
      const step = steps.find((s) => s.title === titles.createOrOpenProject);
      chai.assert.isNotEmpty(step);
      chai.assert.isTrue(step?.condition({} as WholeStatus));
    });

    it("condition: not selected", () => {
      sandbox.stub(condition, "isProjectOpened").returns(true);
      const step = steps.find((s) => s.title === titles.createOrOpenProject);
      chai.assert.isFalse(step?.condition({} as WholeStatus));
    });
  });

  describe(`title: "${titles.summarizeReadme}"`, () => {
    afterEach(() => {
      sandbox.restore();
    });

    it("description", () => {
      const step = steps.find((s) => s.title === titles.summarizeReadme);
      chai.assert.isFalse(
        (step?.description as DescripitionFunc)({
          projectOpened: {
            readmeContent: `
            123456
            # Overview of the AI Assistant Bot template

            This app template is built on top of [Teams AI library](https://aka.ms/teams-ai-library) and [OpenAI Assistants API](https://platform.openai.com/docs/assistants/overview).
            It showcases how to build an intelligent chat bot in Teams capable of helping users accomplish a specific task using natural language right in the Teams conversations, such as solving a math problem.
            
            ## Get started with the AI Assistant Bot template

            > **Prerequisites**`,
          },
        } as WholeStatus).includes("123456")
      );
    });

    it("condition: selected", () => {
      sandbox.stub(condition, "isProjectOpened").returns(true);

      sandbox.stub(condition, "isDidNoActionAfterScaffolded").returns(true);
      sandbox.stub(condition, "isHaveReadMe").returns(true);
      const step = steps.find((s) => s.title === titles.summarizeReadme);
      chai.assert.isNotEmpty(step);
      chai.assert.isTrue(step?.condition({} as WholeStatus));
    });

    it("condition: not selected - no project opened", () => {
      sandbox.stub(condition, "isProjectOpened").returns(false);
      const step = steps.find((s) => s.title === titles.summarizeReadme);
      chai.assert.isFalse(step?.condition({} as WholeStatus));
    });

    it("condition: not selected - prerequisite check failed", () => {
      sandbox.stub(condition, "isProjectOpened").returns(true);

      const step = steps.find((s) => s.title === titles.summarizeReadme);
      chai.assert.isFalse(step?.condition({} as WholeStatus));
    });

    it("condition: not selected - did action before", () => {
      sandbox.stub(condition, "isProjectOpened").returns(true);

      sandbox.stub(condition, "isDidNoActionAfterScaffolded").returns(false);
      const step = steps.find((s) => s.title === titles.summarizeReadme);
      chai.assert.isFalse(step?.condition({} as WholeStatus));
    });

    it("condition: not selected - had no readme content", () => {
      sandbox.stub(condition, "isProjectOpened").returns(true);

      sandbox.stub(condition, "isDidNoActionAfterScaffolded").returns(true);
      sandbox.stub(condition, "isHaveReadMe").returns(false);
      const step = steps.find((s) => s.title === titles.summarizeReadme);
      chai.assert.isFalse(step?.condition({} as WholeStatus));
    });
  });

  describe(`title: "${titles.previewInTestTool}"`, () => {
    afterEach(() => {
      sandbox.restore();
    });

    it("condition: selected", () => {
      sandbox.stub(condition, "isProjectOpened").returns(true);

      sandbox.stub(condition, "isDidNoActionAfterScaffolded").returns(false);
      sandbox.stub(condition, "isDebugSucceededAfterSourceCodeChanged").returns(false);
      sandbox.stub(condition, "canPreviewInTestTool").returns(true);
      const step = steps.find((s) => s.title === titles.previewInTestTool);
      chai.assert.isNotEmpty(step);
      chai.assert.isTrue(step?.condition({} as WholeStatus));
    });

    it("condition: not selected - no project opened", () => {
      sandbox.stub(condition, "isProjectOpened").returns(false);
      const step = steps.find((s) => s.title === titles.previewInTestTool);
      chai.assert.isFalse(step?.condition({} as WholeStatus));
    });

    it("condition: not selected - prerequisite check failed", () => {
      sandbox.stub(condition, "isProjectOpened").returns(true);

      const step = steps.find((s) => s.title === titles.previewInTestTool);
      chai.assert.isFalse(step?.condition({} as WholeStatus));
    });

    it("condition: not selected - did no action before", () => {
      sandbox.stub(condition, "isProjectOpened").returns(true);

      sandbox.stub(condition, "isDidNoActionAfterScaffolded").returns(true);
      const step = steps.find((s) => s.title === titles.previewInTestTool);
      chai.assert.isFalse(step?.condition({} as WholeStatus));
    });

    it("condition: not selected - debug succeed before", () => {
      sandbox.stub(condition, "isProjectOpened").returns(true);

      sandbox.stub(condition, "isDidNoActionAfterScaffolded").returns(false);
      sandbox.stub(condition, "isDebugSucceededAfterSourceCodeChanged").returns(true);
      const step = steps.find((s) => s.title === titles.previewInTestTool);
      chai.assert.isFalse(step?.condition({} as WholeStatus));
    });

    it("condition: not selected - cannot preview in Test Tool", () => {
      sandbox.stub(condition, "isProjectOpened").returns(true);

      sandbox.stub(condition, "isDidNoActionAfterScaffolded").returns(false);
      sandbox.stub(condition, "isDebugSucceededAfterSourceCodeChanged").returns(true);
      sandbox.stub(condition, "canPreviewInTestTool").returns(false);
      const step = steps.find((s) => s.title === titles.previewInTestTool);
      chai.assert.isFalse(step?.condition({} as WholeStatus));
    });
  });

  describe(`title: "${titles.signInM365Account}"`, () => {
    afterEach(() => {
      sandbox.restore();
    });

    it("condition: selected", () => {
      sandbox.stub(condition, "isProjectOpened").returns(true);

      sandbox.stub(condition, "isDidNoActionAfterScaffolded").returns(false);
      sandbox.stub(condition, "isDebugSucceededAfterSourceCodeChanged").returns(false);
      sandbox.stub(condition, "isM365AccountLogin").returns(false);
      const step = steps.find((s) => s.title === titles.signInM365Account);
      chai.assert.isNotEmpty(step);
      chai.assert.isTrue(step?.condition({} as WholeStatus));
    });

    it("condition: not selected - no project opened", () => {
      sandbox.stub(condition, "isProjectOpened").returns(false);
      const step = steps.find((s) => s.title === titles.signInM365Account);
      chai.assert.isFalse(step?.condition({} as WholeStatus));
    });

    it("condition: not selected - prerequisite check failed", () => {
      sandbox.stub(condition, "isProjectOpened").returns(true);

      const step = steps.find((s) => s.title === titles.signInM365Account);
      chai.assert.isFalse(step?.condition({} as WholeStatus));
    });

    it("condition: not selected - did no action before", () => {
      sandbox.stub(condition, "isProjectOpened").returns(true);

      sandbox.stub(condition, "isDidNoActionAfterScaffolded").returns(true);
      const step = steps.find((s) => s.title === titles.signInM365Account);
      chai.assert.isFalse(step?.condition({} as WholeStatus));
    });

    it("condition: not selected - debug succeed before", () => {
      sandbox.stub(condition, "isProjectOpened").returns(true);

      sandbox.stub(condition, "isDidNoActionAfterScaffolded").returns(false);
      sandbox.stub(condition, "isDebugSucceededAfterSourceCodeChanged").returns(true);
      const step = steps.find((s) => s.title === titles.signInM365Account);
      chai.assert.isFalse(step?.condition({} as WholeStatus));
    });

    it("condition: not selected - log into M365 account", () => {
      sandbox.stub(condition, "isProjectOpened").returns(true);

      sandbox.stub(condition, "isDidNoActionAfterScaffolded").returns(false);
      sandbox.stub(condition, "isDebugSucceededAfterSourceCodeChanged").returns(true);
      sandbox.stub(condition, "isM365AccountLogin").returns(true);
      const step = steps.find((s) => s.title === titles.signInM365Account);
      chai.assert.isFalse(step?.condition({} as WholeStatus));
    });
  });

  describe(`title: "${titles.signInM365Account}"`, () => {
    afterEach(() => {
      sandbox.restore();
    });

    it("condition: selected", () => {
      sandbox.stub(condition, "isProjectOpened").returns(true);

      sandbox.stub(condition, "isDidNoActionAfterScaffolded").returns(false);
      sandbox.stub(condition, "isDebugSucceededAfterSourceCodeChanged").returns(false);
      sandbox.stub(condition, "isM365AccountLogin").returns(false);
      const step = steps.find((s) => s.title === titles.signInM365Account);
      chai.assert.isNotEmpty(step);
      chai.assert.isTrue(step?.condition({} as WholeStatus));
    });

    it("condition: not selected - no project opened", () => {
      sandbox.stub(condition, "isProjectOpened").returns(false);
      const step = steps.find((s) => s.title === titles.signInM365Account);
      chai.assert.isFalse(step?.condition({} as WholeStatus));
    });

    it("condition: not selected - prerequisite check failed", () => {
      sandbox.stub(condition, "isProjectOpened").returns(true);

      const step = steps.find((s) => s.title === titles.signInM365Account);
      chai.assert.isFalse(step?.condition({} as WholeStatus));
    });

    it("condition: not selected - did no action before", () => {
      sandbox.stub(condition, "isProjectOpened").returns(true);

      sandbox.stub(condition, "isDidNoActionAfterScaffolded").returns(true);
      const step = steps.find((s) => s.title === titles.signInM365Account);
      chai.assert.isFalse(step?.condition({} as WholeStatus));
    });

    it("condition: not selected - debug succeed before", () => {
      sandbox.stub(condition, "isProjectOpened").returns(true);

      sandbox.stub(condition, "isDidNoActionAfterScaffolded").returns(false);
      sandbox.stub(condition, "isDebugSucceededAfterSourceCodeChanged").returns(true);
      const step = steps.find((s) => s.title === titles.signInM365Account);
      chai.assert.isFalse(step?.condition({} as WholeStatus));
    });

    it("condition: not selected - log into M365 account", () => {
      sandbox.stub(condition, "isProjectOpened").returns(true);

      sandbox.stub(condition, "isDidNoActionAfterScaffolded").returns(false);
      sandbox.stub(condition, "isDebugSucceededAfterSourceCodeChanged").returns(true);
      sandbox.stub(condition, "isM365AccountLogin").returns(true);
      const step = steps.find((s) => s.title === titles.signInM365Account);
      chai.assert.isFalse(step?.condition({} as WholeStatus));
    });
  });

  describe(`title: "${titles.previewInTeams}"`, () => {
    afterEach(() => {
      sandbox.restore();
    });

    it("condition: selected", () => {
      sandbox.stub(condition, "isProjectOpened").returns(true);

      sandbox.stub(condition, "isDidNoActionAfterScaffolded").returns(false);
      sandbox.stub(condition, "isDebugSucceededAfterSourceCodeChanged").returns(false);
      sandbox.stub(condition, "isM365AccountLogin").returns(true);
      const step = steps.find((s) => s.title === titles.previewInTeams);
      chai.assert.isNotEmpty(step);
      chai.assert.isTrue(step?.condition({} as WholeStatus));
    });

    it("condition: not selected - no project opened", () => {
      sandbox.stub(condition, "isProjectOpened").returns(false);
      const step = steps.find((s) => s.title === titles.previewInTeams);
      chai.assert.isFalse(step?.condition({} as WholeStatus));
    });

    it("condition: not selected - prerequisite check failed", () => {
      sandbox.stub(condition, "isProjectOpened").returns(true);

      const step = steps.find((s) => s.title === titles.previewInTeams);
      chai.assert.isFalse(step?.condition({} as WholeStatus));
    });

    it("condition: not selected - did no action before", () => {
      sandbox.stub(condition, "isProjectOpened").returns(true);

      sandbox.stub(condition, "isDidNoActionAfterScaffolded").returns(true);
      const step = steps.find((s) => s.title === titles.previewInTeams);
      chai.assert.isFalse(step?.condition({} as WholeStatus));
    });

    it("condition: not selected - debug succeed before", () => {
      sandbox.stub(condition, "isProjectOpened").returns(true);

      sandbox.stub(condition, "isDidNoActionAfterScaffolded").returns(false);
      sandbox.stub(condition, "isDebugSucceededAfterSourceCodeChanged").returns(true);
      const step = steps.find((s) => s.title === titles.previewInTeams);
      chai.assert.isFalse(step?.condition({} as WholeStatus));
    });

    it("condition: not selected - not log into M365 account", () => {
      sandbox.stub(condition, "isProjectOpened").returns(true);

      sandbox.stub(condition, "isDidNoActionAfterScaffolded").returns(false);
      sandbox.stub(condition, "isDebugSucceededAfterSourceCodeChanged").returns(true);
      sandbox.stub(condition, "isM365AccountLogin").returns(false);
      const step = steps.find((s) => s.title === titles.previewInTeams);
      chai.assert.isFalse(step?.condition({} as WholeStatus));
    });
  });

  describe(`title: "${titles.howToExtend}"`, () => {
    afterEach(() => {
      sandbox.restore();
    });

    it("description", () => {
      const step = steps.find((s) => s.title === titles.howToExtend);
      chai.assert.isTrue(
        (step?.description as DescripitionFunc)({
          projectOpened: {
            readmeContent: `
            ### Run Teams Bot locally

            ## What's included in the template

            ## Extend the AI Assistant Bot template with more AI capabilities`,
          },
        } as WholeStatus).includes("Extend the AI Assistant Bot template with more AI capabilities")
      );
    });

    it("condition: selected", () => {
      sandbox.stub(condition, "isProjectOpened").returns(true);

      sandbox.stub(condition, "isDidNoActionAfterScaffolded").returns(false);
      sandbox.stub(condition, "isDebugSucceededAfterSourceCodeChanged").returns(true);
      sandbox.stub(condition, "isHaveReadMe").returns(true);
      const step = steps.find((s) => s.title === titles.howToExtend);
      chai.assert.isNotEmpty(step);
      chai.assert.isTrue(step?.condition({} as WholeStatus));
    });

    it("condition: not selected - no project opened", () => {
      sandbox.stub(condition, "isProjectOpened").returns(false);
      const step = steps.find((s) => s.title === titles.howToExtend);
      chai.assert.isFalse(step?.condition({} as WholeStatus));
    });

    it("condition: not selected - prerequisite check failed", () => {
      sandbox.stub(condition, "isProjectOpened").returns(true);

      const step = steps.find((s) => s.title === titles.howToExtend);
      chai.assert.isFalse(step?.condition({} as WholeStatus));
    });

    it("condition: not selected - did no action before", () => {
      sandbox.stub(condition, "isProjectOpened").returns(true);

      sandbox.stub(condition, "isDidNoActionAfterScaffolded").returns(true);
      const step = steps.find((s) => s.title === titles.howToExtend);
      chai.assert.isFalse(step?.condition({} as WholeStatus));
    });

    it("condition: not selected - debug failed before", () => {
      sandbox.stub(condition, "isProjectOpened").returns(true);

      sandbox.stub(condition, "isDidNoActionAfterScaffolded").returns(false);
      sandbox.stub(condition, "isDebugSucceededAfterSourceCodeChanged").returns(false);
      const step = steps.find((s) => s.title === titles.howToExtend);
      chai.assert.isFalse(step?.condition({} as WholeStatus));
    });

    it("condition: not selected - had no readme content", () => {
      sandbox.stub(condition, "isProjectOpened").returns(true);

      sandbox.stub(condition, "isDidNoActionAfterScaffolded").returns(false);
      sandbox.stub(condition, "isDebugSucceededAfterSourceCodeChanged").returns(true);
      sandbox.stub(condition, "isHaveReadMe").returns(false);
      const step = steps.find((s) => s.title === titles.howToExtend);
      chai.assert.isFalse(step?.condition({} as WholeStatus));
    });
  });

  describe(`title: "${titles.ciCd}"`, () => {
    afterEach(() => {
      sandbox.restore();
    });

    it("condition: selected", () => {
      sandbox.stub(condition, "isProjectOpened").returns(true);

      sandbox.stub(condition, "isDidNoActionAfterScaffolded").returns(false);
      sandbox.stub(condition, "isDebugSucceededAfterSourceCodeChanged").returns(true);
      const step = steps.find((s) => s.title === titles.ciCd);
      chai.assert.isNotEmpty(step);
      chai.assert.isTrue(step?.condition({} as WholeStatus));
    });

    it("condition: not selected - no project opened", () => {
      sandbox.stub(condition, "isProjectOpened").returns(false);
      const step = steps.find((s) => s.title === titles.ciCd);
      chai.assert.isFalse(step?.condition({} as WholeStatus));
    });

    it("condition: not selected - prerequisite check failed", () => {
      sandbox.stub(condition, "isProjectOpened").returns(true);

      const step = steps.find((s) => s.title === titles.ciCd);
      chai.assert.isFalse(step?.condition({} as WholeStatus));
    });

    it("condition: not selected - did no action before", () => {
      sandbox.stub(condition, "isProjectOpened").returns(true);

      sandbox.stub(condition, "isDidNoActionAfterScaffolded").returns(true);
      const step = steps.find((s) => s.title === titles.ciCd);
      chai.assert.isFalse(step?.condition({} as WholeStatus));
    });

    it("condition: not selected - debug failed before", () => {
      sandbox.stub(condition, "isProjectOpened").returns(true);

      sandbox.stub(condition, "isDidNoActionAfterScaffolded").returns(false);
      sandbox.stub(condition, "isDebugSucceededAfterSourceCodeChanged").returns(false);
      const step = steps.find((s) => s.title === titles.ciCd);
      chai.assert.isFalse(step?.condition({} as WholeStatus));
    });
  });

  describe(`title: "${titles.azureAccount}"`, () => {
    afterEach(() => {
      sandbox.restore();
    });

    it("condition: selected", () => {
      sandbox.stub(condition, "isProjectOpened").returns(true);

      sandbox.stub(condition, "isDidNoActionAfterScaffolded").returns(false);
      sandbox.stub(condition, "isDebugSucceededAfterSourceCodeChanged").returns(true);
      sandbox.stub(condition, "isProvisionedSucceededAfterInfraCodeChanged").returns(false);
      sandbox.stub(condition, "isAzureAccountLogin").returns(false);
      const step = steps.find((s) => s.title === titles.azureAccount);
      chai.assert.isNotEmpty(step);
      chai.assert.isTrue(step?.condition({} as WholeStatus));
    });

    it("condition: not selected - no project opened", () => {
      sandbox.stub(condition, "isProjectOpened").returns(false);
      const step = steps.find((s) => s.title === titles.azureAccount);
      chai.assert.isFalse(step?.condition({} as WholeStatus));
    });

    it("condition: not selected - prerequisite check failed", () => {
      sandbox.stub(condition, "isProjectOpened").returns(true);

      const step = steps.find((s) => s.title === titles.azureAccount);
      chai.assert.isFalse(step?.condition({} as WholeStatus));
    });

    it("condition: not selected - did no action before", () => {
      sandbox.stub(condition, "isProjectOpened").returns(true);

      sandbox.stub(condition, "isDidNoActionAfterScaffolded").returns(true);
      const step = steps.find((s) => s.title === titles.azureAccount);
      chai.assert.isFalse(step?.condition({} as WholeStatus));
    });

    it("condition: not selected - debug failed before", () => {
      sandbox.stub(condition, "isProjectOpened").returns(true);

      sandbox.stub(condition, "isDidNoActionAfterScaffolded").returns(false);
      sandbox.stub(condition, "isDebugSucceededAfterSourceCodeChanged").returns(false);
      const step = steps.find((s) => s.title === titles.azureAccount);
      chai.assert.isFalse(step?.condition({} as WholeStatus));
    });

    it("condition: not selected - provision succeeded before", () => {
      sandbox.stub(condition, "isProjectOpened").returns(true);

      sandbox.stub(condition, "isDidNoActionAfterScaffolded").returns(false);
      sandbox.stub(condition, "isDebugSucceededAfterSourceCodeChanged").returns(true);
      sandbox.stub(condition, "isProvisionedSucceededAfterInfraCodeChanged").returns(true);
      const step = steps.find((s) => s.title === titles.azureAccount);
      chai.assert.isFalse(step?.condition({} as WholeStatus));
    });

    it("condition: not selected - not log into Azure account", () => {
      sandbox.stub(condition, "isProjectOpened").returns(true);

      sandbox.stub(condition, "isDidNoActionAfterScaffolded").returns(false);
      sandbox.stub(condition, "isDebugSucceededAfterSourceCodeChanged").returns(true);
      sandbox.stub(condition, "isProvisionedSucceededAfterInfraCodeChanged").returns(false);
      sandbox.stub(condition, "isAzureAccountLogin").returns(true);
      const step = steps.find((s) => s.title === titles.azureAccount);
      chai.assert.isFalse(step?.condition({} as WholeStatus));
    });
  });

  describe(`title: "${titles.provision}"`, () => {
    afterEach(() => {
      sandbox.restore();
    });

    it("condition: selected", () => {
      sandbox.stub(condition, "isProjectOpened").returns(true);

      sandbox.stub(condition, "isDidNoActionAfterScaffolded").returns(false);
      sandbox.stub(condition, "isDebugSucceededAfterSourceCodeChanged").returns(true);
      sandbox.stub(condition, "isProvisionedSucceededAfterInfraCodeChanged").returns(false);
      sandbox.stub(condition, "isAzureAccountLogin").returns(true);
      const step = steps.find((s) => s.title === titles.provision);
      chai.assert.isNotEmpty(step);
      chai.assert.isTrue(step?.condition({} as WholeStatus));
    });

    it("condition: not selected - no project opened", () => {
      sandbox.stub(condition, "isProjectOpened").returns(false);
      const step = steps.find((s) => s.title === titles.provision);
      chai.assert.isFalse(step?.condition({} as WholeStatus));
    });

    it("condition: not selected - prerequisite check failed", () => {
      sandbox.stub(condition, "isProjectOpened").returns(true);

      const step = steps.find((s) => s.title === titles.provision);
      chai.assert.isFalse(step?.condition({} as WholeStatus));
    });

    it("condition: not selected - did no action before", () => {
      sandbox.stub(condition, "isProjectOpened").returns(true);

      sandbox.stub(condition, "isDidNoActionAfterScaffolded").returns(true);
      const step = steps.find((s) => s.title === titles.provision);
      chai.assert.isFalse(step?.condition({} as WholeStatus));
    });

    it("condition: not selected - debug failed before", () => {
      sandbox.stub(condition, "isProjectOpened").returns(true);

      sandbox.stub(condition, "isDidNoActionAfterScaffolded").returns(false);
      sandbox.stub(condition, "isDebugSucceededAfterSourceCodeChanged").returns(false);
      const step = steps.find((s) => s.title === titles.provision);
      chai.assert.isFalse(step?.condition({} as WholeStatus));
    });

    it("condition: not selected - provision succeeded before", () => {
      sandbox.stub(condition, "isProjectOpened").returns(true);

      sandbox.stub(condition, "isDidNoActionAfterScaffolded").returns(false);
      sandbox.stub(condition, "isDebugSucceededAfterSourceCodeChanged").returns(true);
      sandbox.stub(condition, "isProvisionedSucceededAfterInfraCodeChanged").returns(true);
      const step = steps.find((s) => s.title === titles.provision);
      chai.assert.isFalse(step?.condition({} as WholeStatus));
    });

    it("condition: not selected - not log into Azure Account", () => {
      sandbox.stub(condition, "isProjectOpened").returns(true);

      sandbox.stub(condition, "isDidNoActionAfterScaffolded").returns(false);
      sandbox.stub(condition, "isDebugSucceededAfterSourceCodeChanged").returns(true);
      sandbox.stub(condition, "isProvisionedSucceededAfterInfraCodeChanged").returns(false);
      sandbox.stub(condition, "isAzureAccountLogin").returns(false);
      const step = steps.find((s) => s.title === titles.provision);
      chai.assert.isFalse(step?.condition({} as WholeStatus));
    });
  });

  describe(`title: "${titles.deploy}"`, () => {
    afterEach(() => {
      sandbox.restore();
    });

    it("condition: selected", () => {
      sandbox.stub(condition, "isProjectOpened").returns(true);

      sandbox.stub(condition, "isDidNoActionAfterScaffolded").returns(false);
      sandbox.stub(condition, "isDebugSucceededAfterSourceCodeChanged").returns(true);
      sandbox.stub(condition, "isProvisionedSucceededAfterInfraCodeChanged").returns(true);
      sandbox.stub(condition, "isDeployedAfterSourceCodeChanged").returns(false);
      const step = steps.find((s) => s.title === titles.deploy);
      chai.assert.isNotEmpty(step);
      chai.assert.isTrue(step?.condition({} as WholeStatus));
    });

    it("condition: not selected - no project opened", () => {
      sandbox.stub(condition, "isProjectOpened").returns(false);
      const step = steps.find((s) => s.title === titles.deploy);
      chai.assert.isFalse(step?.condition({} as WholeStatus));
    });

    it("condition: not selected - prerequisite check failed", () => {
      sandbox.stub(condition, "isProjectOpened").returns(true);

      const step = steps.find((s) => s.title === titles.deploy);
      chai.assert.isFalse(step?.condition({} as WholeStatus));
    });

    it("condition: not selected - did no action before", () => {
      sandbox.stub(condition, "isProjectOpened").returns(true);

      sandbox.stub(condition, "isDidNoActionAfterScaffolded").returns(true);
      const step = steps.find((s) => s.title === titles.deploy);
      chai.assert.isFalse(step?.condition({} as WholeStatus));
    });

    it("condition: not selected - debug failed before", () => {
      sandbox.stub(condition, "isProjectOpened").returns(true);

      sandbox.stub(condition, "isDidNoActionAfterScaffolded").returns(false);
      sandbox.stub(condition, "isDebugSucceededAfterSourceCodeChanged").returns(false);
      const step = steps.find((s) => s.title === titles.deploy);
      chai.assert.isFalse(step?.condition({} as WholeStatus));
    });

    it("condition: not selected - provision failed before", () => {
      sandbox.stub(condition, "isProjectOpened").returns(true);

      sandbox.stub(condition, "isDidNoActionAfterScaffolded").returns(false);
      sandbox.stub(condition, "isDebugSucceededAfterSourceCodeChanged").returns(true);
      sandbox.stub(condition, "isProvisionedSucceededAfterInfraCodeChanged").returns(false);
      const step = steps.find((s) => s.title === titles.deploy);
      chai.assert.isFalse(step?.condition({} as WholeStatus));
    });

    it("condition: not selected - deploy succeeded before", () => {
      sandbox.stub(condition, "isProjectOpened").returns(true);

      sandbox.stub(condition, "isDidNoActionAfterScaffolded").returns(false);
      sandbox.stub(condition, "isDebugSucceededAfterSourceCodeChanged").returns(true);
      sandbox.stub(condition, "isProvisionedSucceededAfterInfraCodeChanged").returns(true);
      sandbox.stub(condition, "isDeployedAfterSourceCodeChanged").returns(true);
      const step = steps.find((s) => s.title === titles.deploy);
      chai.assert.isFalse(step?.condition({} as WholeStatus));
    });
  });

  describe(`title: "${titles.publish}"`, () => {
    afterEach(() => {
      sandbox.restore();
    });

    it("condition: selected", () => {
      sandbox.stub(condition, "isProjectOpened").returns(true);

      sandbox.stub(condition, "isDidNoActionAfterScaffolded").returns(false);
      sandbox.stub(condition, "isDebugSucceededAfterSourceCodeChanged").returns(true);
      sandbox.stub(condition, "isProvisionedSucceededAfterInfraCodeChanged").returns(true);
      sandbox.stub(condition, "isDeployedAfterSourceCodeChanged").returns(true);
      sandbox.stub(condition, "isPublishedSucceededBefore").returns(false);
      const step = steps.find((s) => s.title === titles.publish);
      chai.assert.isNotEmpty(step);
      chai.assert.isTrue(step?.condition({} as WholeStatus));
    });

    it("condition: not selected - no project opened", () => {
      sandbox.stub(condition, "isProjectOpened").returns(false);
      const step = steps.find((s) => s.title === titles.publish);
      chai.assert.isFalse(step?.condition({} as WholeStatus));
    });

    it("condition: not selected - prerequisite check failed", () => {
      sandbox.stub(condition, "isProjectOpened").returns(true);

      const step = steps.find((s) => s.title === titles.publish);
      chai.assert.isFalse(step?.condition({} as WholeStatus));
    });

    it("condition: not selected - did no action before", () => {
      sandbox.stub(condition, "isProjectOpened").returns(true);

      sandbox.stub(condition, "isDidNoActionAfterScaffolded").returns(true);
      const step = steps.find((s) => s.title === titles.publish);
      chai.assert.isFalse(step?.condition({} as WholeStatus));
    });

    it("condition: not selected - debug failed before", () => {
      sandbox.stub(condition, "isProjectOpened").returns(true);

      sandbox.stub(condition, "isDidNoActionAfterScaffolded").returns(false);
      sandbox.stub(condition, "isDebugSucceededAfterSourceCodeChanged").returns(false);
      const step = steps.find((s) => s.title === titles.publish);
      chai.assert.isFalse(step?.condition({} as WholeStatus));
    });

    it("condition: not selected - provision failed before", () => {
      sandbox.stub(condition, "isProjectOpened").returns(true);

      sandbox.stub(condition, "isDidNoActionAfterScaffolded").returns(false);
      sandbox.stub(condition, "isDebugSucceededAfterSourceCodeChanged").returns(true);
      sandbox.stub(condition, "isProvisionedSucceededAfterInfraCodeChanged").returns(false);
      const step = steps.find((s) => s.title === titles.publish);
      chai.assert.isFalse(step?.condition({} as WholeStatus));
    });

    it("condition: not selected - deploy failed before", () => {
      sandbox.stub(condition, "isProjectOpened").returns(true);

      sandbox.stub(condition, "isDidNoActionAfterScaffolded").returns(false);
      sandbox.stub(condition, "isDebugSucceededAfterSourceCodeChanged").returns(true);
      sandbox.stub(condition, "isProvisionedSucceededAfterInfraCodeChanged").returns(true);
      sandbox.stub(condition, "isDeployedAfterSourceCodeChanged").returns(false);
      const step = steps.find((s) => s.title === titles.publish);
      chai.assert.isFalse(step?.condition({} as WholeStatus));
    });

    it("condition: not selected - published before", () => {
      sandbox.stub(condition, "isProjectOpened").returns(true);

      sandbox.stub(condition, "isDidNoActionAfterScaffolded").returns(false);
      sandbox.stub(condition, "isDebugSucceededAfterSourceCodeChanged").returns(true);
      sandbox.stub(condition, "isProvisionedSucceededAfterInfraCodeChanged").returns(true);
      sandbox.stub(condition, "isDeployedAfterSourceCodeChanged").returns(true);
      sandbox.stub(condition, "isPublishedSucceededBefore").returns(true);
      const step = steps.find((s) => s.title === titles.publish);
      chai.assert.isFalse(step?.condition({} as WholeStatus));
    });
  });

  describe(`title: "${titles.previewRemotely}"`, () => {
    afterEach(() => {
      sandbox.restore();
    });

    it("condition: selected", () => {
      sandbox.stub(condition, "isProjectOpened").returns(true);

      sandbox.stub(condition, "isDidNoActionAfterScaffolded").returns(false);
      sandbox.stub(condition, "isDebugSucceededAfterSourceCodeChanged").returns(true);
      sandbox.stub(condition, "isProvisionedSucceededAfterInfraCodeChanged").returns(true);
      sandbox.stub(condition, "isDeployedAfterSourceCodeChanged").returns(true);
      const step = steps.find((s) => s.title === titles.previewRemotely);
      chai.assert.isNotEmpty(step);
      chai.assert.isTrue(step?.condition({} as WholeStatus));
    });

    it("condition: not selected - no project opened", () => {
      sandbox.stub(condition, "isProjectOpened").returns(false);
      const step = steps.find((s) => s.title === titles.previewRemotely);
      chai.assert.isFalse(step?.condition({} as WholeStatus));
    });

    it("condition: not selected - prerequisite check failed", () => {
      sandbox.stub(condition, "isProjectOpened").returns(true);

      const step = steps.find((s) => s.title === titles.previewRemotely);
      chai.assert.isFalse(step?.condition({} as WholeStatus));
    });

    it("condition: not selected - did no action before", () => {
      sandbox.stub(condition, "isProjectOpened").returns(true);

      sandbox.stub(condition, "isDidNoActionAfterScaffolded").returns(true);
      const step = steps.find((s) => s.title === titles.previewRemotely);
      chai.assert.isFalse(step?.condition({} as WholeStatus));
    });

    it("condition: not selected - debug failed before", () => {
      sandbox.stub(condition, "isProjectOpened").returns(true);

      sandbox.stub(condition, "isDidNoActionAfterScaffolded").returns(false);
      sandbox.stub(condition, "isDebugSucceededAfterSourceCodeChanged").returns(false);
      const step = steps.find((s) => s.title === titles.previewRemotely);
      chai.assert.isFalse(step?.condition({} as WholeStatus));
    });

    it("condition: not selected - provision failed before", () => {
      sandbox.stub(condition, "isProjectOpened").returns(true);

      sandbox.stub(condition, "isDidNoActionAfterScaffolded").returns(false);
      sandbox.stub(condition, "isDebugSucceededAfterSourceCodeChanged").returns(true);
      sandbox.stub(condition, "isProvisionedSucceededAfterInfraCodeChanged").returns(false);
      const step = steps.find((s) => s.title === titles.previewRemotely);
      chai.assert.isFalse(step?.condition({} as WholeStatus));
    });

    it("condition: not selected - deploy failed before", () => {
      sandbox.stub(condition, "isProjectOpened").returns(true);

      sandbox.stub(condition, "isDidNoActionAfterScaffolded").returns(false);
      sandbox.stub(condition, "isDebugSucceededAfterSourceCodeChanged").returns(true);
      sandbox.stub(condition, "isProvisionedSucceededAfterInfraCodeChanged").returns(true);
      sandbox.stub(condition, "isDeployedAfterSourceCodeChanged").returns(false);
      const step = steps.find((s) => s.title === titles.previewRemotely);
      chai.assert.isFalse(step?.condition({} as WholeStatus));
    });
  });
});
