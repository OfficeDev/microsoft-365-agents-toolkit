import * as chai from "chai";
import * as sinon from "sinon";
import { officeSteps } from "../../../../src/officeChat/commands/nextStep/officeSteps";
import * as condition from "../../../../src/officeChat/commands/nextStep/condition";
import { OfficeWholeStatus } from "../../../../src/officeChat/commands/nextStep/types";

describe("office steps", () => {
  const sandbox = sinon.createSandbox();
  const steps = officeSteps();

  describe('title: "Create a New Project"', () => {
    afterEach(() => {
      sandbox.restore();
    });

    it("condition: selected", () => {
      sandbox.stub(condition, "isProjectOpened").returns(false);
      const newProject = steps.find((s) => s.title === "Create a New Project");
      chai.assert.isNotEmpty(newProject);
      chai.assert.isTrue(newProject?.condition({} as OfficeWholeStatus));
    });

    it("condition: not selected", () => {
      sandbox.stub(condition, "isProjectOpened").returns(true);
      const newProject = steps.find((s) => s.title === "Create a New Project");
      chai.assert.isFalse(newProject?.condition({} as OfficeWholeStatus));
    });
  });

  describe('title: "Check Prerequisites"', () => {
    afterEach(() => {
      sandbox.restore();
    });

    it("condition: selected", () => {
      sandbox.stub(condition, "isProjectOpened").returns(true);
      sandbox.stub(condition, "isNodeInstalled").returns(false);
      const step = steps.find((s) => s.title === "Check Prerequisites");
      chai.assert.isNotEmpty(step);
      chai.assert.isTrue(step?.condition({} as OfficeWholeStatus));
    });

    it("condition: not selected - no project opened", () => {
      sandbox.stub(condition, "isProjectOpened").returns(false);
      const step = steps.find((s) => s.title === "Check Prerequisites");
      chai.assert.isFalse(step?.condition({} as OfficeWholeStatus));
    });

    it("condition: not selected - prerequisite check succeeded", () => {
      sandbox.stub(condition, "isProjectOpened").returns(true);
      sandbox.stub(condition, "isNodeInstalled").returns(true);
      const step = steps.find((s) => s.title === "Check Prerequisites");
      chai.assert.isFalse(step?.condition({} as OfficeWholeStatus));
    });
  });

  describe("Install Dependencies", () => {
    afterEach(() => {
      sandbox.restore();
    });

    it("condition: selected - project opened, did action after scaffolded, dependencies not installed", () => {
      sandbox.stub(condition, "isProjectOpened").returns(true);
      sandbox.stub(condition, "isNodeInstalled").returns(true);
      sandbox.stub(condition, "isDependenciesInstalled").returns(false);

      const step = steps.find((s) => s.title === "Install Dependencies");
      chai.assert.isTrue(step?.condition({} as OfficeWholeStatus));
    });

    it("condition: not selected - project not opened", () => {
      sandbox.stub(condition, "isProjectOpened").returns(false);

      const step = steps.find((s) => s.title === "Install Dependencies");
      chai.assert.isFalse(step?.condition({} as OfficeWholeStatus));
    });

    it("condition: not selected - prerequisite check failed", () => {
      sandbox.stub(condition, "isProjectOpened").returns(true);
      sandbox.stub(condition, "isNodeInstalled").returns(false);

      const step = steps.find((s) => s.title === "Install Dependencies");
      chai.assert.isFalse(step?.condition({} as OfficeWholeStatus));
    });

    it("condition: not selected - dependencies installed", () => {
      sandbox.stub(condition, "isProjectOpened").returns(true);
      sandbox.stub(condition, "isNodeInstalled").returns(true);
      sandbox.stub(condition, "isDependenciesInstalled").returns(true);

      const step = steps.find((s) => s.title === "Install Dependencies");
      chai.assert.isFalse(step?.condition({} as OfficeWholeStatus));
    });
  });

  describe("Preview in Local Environment", () => {
    afterEach(() => {
      sandbox.restore();
    });

    it("condition: selected - project opened, node installed, dependencies installed, can preview in local env, debug not succeeded after source code changed", () => {
      sandbox.stub(condition, "isProjectOpened").returns(true);
      sandbox.stub(condition, "isNodeInstalled").returns(true);
      sandbox.stub(condition, "isDependenciesInstalled").returns(true);
      sandbox.stub(condition, "canOfficeAddInPreviewInLocalEnv").returns(true);
      sandbox.stub(condition, "isDebugSucceededAfterSourceCodeChanged").returns(false);

      const step = steps.find((s) => s.title === "Preview in Local Environment");
      chai.assert.isTrue(step?.condition({} as OfficeWholeStatus));
    });

    it("condition: not selected - project not opened", () => {
      sandbox.stub(condition, "isProjectOpened").returns(false);

      const step = steps.find((s) => s.title === "Preview in Local Environment");
      chai.assert.isFalse(step?.condition({} as OfficeWholeStatus));
    });

    it("condition: not selected - node not installed", () => {
      sandbox.stub(condition, "isProjectOpened").returns(true);
      sandbox.stub(condition, "isNodeInstalled").returns(false);

      const step = steps.find((s) => s.title === "Preview in Local Environment");
      chai.assert.isFalse(step?.condition({} as OfficeWholeStatus));
    });

    it("condition: not selected - dependencies not installed", () => {
      sandbox.stub(condition, "isProjectOpened").returns(true);
      sandbox.stub(condition, "isNodeInstalled").returns(true);
      sandbox.stub(condition, "isDependenciesInstalled").returns(false);

      const step = steps.find((s) => s.title === "Preview in Local Environment");
      chai.assert.isFalse(step?.condition({} as OfficeWholeStatus));
    });

    it("condition: not selected - cannot preview in local env", () => {
      sandbox.stub(condition, "isProjectOpened").returns(true);
      sandbox.stub(condition, "isNodeInstalled").returns(true);
      sandbox.stub(condition, "isDependenciesInstalled").returns(true);
      sandbox.stub(condition, "canOfficeAddInPreviewInLocalEnv").returns(false);

      const step = steps.find((s) => s.title === "Preview in Local Environment");
      chai.assert.isFalse(step?.condition({} as OfficeWholeStatus));
    });

    it("condition: not selected - debug succeeded after source code changed", () => {
      sandbox.stub(condition, "isProjectOpened").returns(true);
      sandbox.stub(condition, "isNodeInstalled").returns(true);
      sandbox.stub(condition, "isDependenciesInstalled").returns(true);
      sandbox.stub(condition, "canOfficeAddInPreviewInLocalEnv").returns(true);
      sandbox.stub(condition, "isDebugSucceededAfterSourceCodeChanged").returns(true);

      const step = steps.find((s) => s.title === "Preview in Local Environment");
      chai.assert.isFalse(step?.condition({} as OfficeWholeStatus));
    });
  });

  describe("Code Gen & Deploy", () => {
    afterEach(() => {
      sandbox.restore();
    });

    it("condition: selected - project opened, node installed, dependencies installed, debug succeeded after source code changed", () => {
      sandbox.stub(condition, "isProjectOpened").returns(true);
      sandbox.stub(condition, "isNodeInstalled").returns(true);
      sandbox.stub(condition, "isDependenciesInstalled").returns(true);
      sandbox.stub(condition, "isDebugSucceededAfterSourceCodeChanged").returns(true);

      const step = steps.filter(
        (s) => s.title === "Code Generation" || s.title === "Deploy or Publish"
      );
      chai.assert.isTrue(step?.[0]?.condition({} as OfficeWholeStatus));
      chai.assert.isTrue(step?.[1]?.condition({} as OfficeWholeStatus));
    });

    it("condition: not selected - project not opened", () => {
      sandbox.stub(condition, "isProjectOpened").returns(false);

      const step = steps.filter(
        (s) => s.title === "Code Generation" || s.title === "Deploy or Publish"
      );
      chai.assert.isFalse(step?.[0]?.condition({} as OfficeWholeStatus));
      chai.assert.isFalse(step?.[1]?.condition({} as OfficeWholeStatus));
    });

    it("condition: not selected - node not installed", () => {
      sandbox.stub(condition, "isProjectOpened").returns(true);
      sandbox.stub(condition, "isNodeInstalled").returns(false);

      const step = steps.filter(
        (s) => s.title === "Code Generation" || s.title === "Deploy or Publish"
      );
      chai.assert.isFalse(step?.[0]?.condition({} as OfficeWholeStatus));
      chai.assert.isFalse(step?.[1]?.condition({} as OfficeWholeStatus));
    });

    it("condition: not selected - dependencies not installed", () => {
      sandbox.stub(condition, "isProjectOpened").returns(true);
      sandbox.stub(condition, "isNodeInstalled").returns(true);
      sandbox.stub(condition, "isDependenciesInstalled").returns(false);

      const step = steps.filter(
        (s) => s.title === "Code Generation" || s.title === "Deploy or Publish"
      );
      chai.assert.isFalse(step?.[0]?.condition({} as OfficeWholeStatus));
      chai.assert.isFalse(step?.[1]?.condition({} as OfficeWholeStatus));
    });

    it("condition: not selected - debug not succeeded after source code changed", () => {
      sandbox.stub(condition, "isProjectOpened").returns(true);
      sandbox.stub(condition, "isNodeInstalled").returns(true);
      sandbox.stub(condition, "isDependenciesInstalled").returns(true);
      sandbox.stub(condition, "isDebugSucceededAfterSourceCodeChanged").returns(false);

      const step = steps.filter(
        (s) => s.title === "Code Generation" || s.title === "Deploy or Publish"
      );
      chai.assert.isFalse(step?.[0]?.condition({} as OfficeWholeStatus));
      chai.assert.isFalse(step?.[1]?.condition({} as OfficeWholeStatus));
    });
  });
});
