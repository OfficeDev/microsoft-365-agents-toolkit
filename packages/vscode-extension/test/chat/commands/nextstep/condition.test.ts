import * as chai from "chai";
import chaiPromised from "chai-as-promised";
import * as condition from "../../../../src/chat/commands/nextstep/condition";
import { WholeStatus } from "../../../../src/chat/commands/nextstep/types";
import { CommandKey } from "../../../../src/constants";
import { emptyProjectStatus } from "../../../../src/utils/projectStatusUtils";

chai.use(chaiPromised);

describe("chat nextstep conditions", () => {
  it("isFirstInstalled", () => {
    chai.assert.isTrue(
      condition.isFirstInstalled({
        machineStatus: {
          firstInstalled: true,
        },
      } as WholeStatus)
    );
  });

  it("isProjectOpened", () => {
    chai.assert.isTrue(
      condition.isProjectOpened({
        projectOpened: {},
      } as WholeStatus)
    );
    chai.assert.isFalse(condition.isProjectOpened({} as WholeStatus));
  });

  describe("isDidNoActionAfterScaffolded", () => {
    it("no opened project", () => {
      chai.assert.isTrue(condition.isDidNoActionAfterScaffolded({} as WholeStatus));
    });

    it("action status is empty", () => {
      chai.assert.isTrue(
        condition.isDidNoActionAfterScaffolded({
          projectOpened: {
            actionStatus: emptyProjectStatus(),
          },
        } as WholeStatus)
      );
    });

    it("some action is done", () => {
      chai.assert.isFalse(
        condition.isDidNoActionAfterScaffolded({
          projectOpened: {
            actionStatus: {
              ...emptyProjectStatus(),
              [CommandKey.Provision]: { result: "success", time: new Date() },
            },
          },
        } as WholeStatus)
      );
    });

    it("some action is failed", () => {
      chai.assert.isFalse(
        condition.isDidNoActionAfterScaffolded({
          projectOpened: {
            actionStatus: {
              ...emptyProjectStatus(),
              [CommandKey.Provision]: { result: "fail", time: new Date() },
            },
          },
        } as WholeStatus)
      );
    });
  });

  describe("isDebugSucceededAfterSourceCodeChanged", () => {
    it("no opened project", () => {
      chai.assert.isFalse(condition.isDebugSucceededAfterSourceCodeChanged({} as WholeStatus));
    });

    it("local debug not run before", () => {
      chai.assert.isFalse(
        condition.isDebugSucceededAfterSourceCodeChanged({
          projectOpened: {
            actionStatus: {
              [CommandKey.LocalDebug]: { result: "no run", time: new Date() },
            },
          },
        } as WholeStatus)
      );
    });

    it("local debug failed before", () => {
      chai.assert.isFalse(
        condition.isDebugSucceededAfterSourceCodeChanged({
          projectOpened: {
            actionStatus: {
              [CommandKey.LocalDebug]: { result: "fail", time: new Date() },
            },
          },
        } as WholeStatus)
      );
    });

    it("local debug succeeded before but out of date", () => {
      chai.assert.isFalse(
        condition.isDebugSucceededAfterSourceCodeChanged({
          projectOpened: {
            actionStatus: {
              [CommandKey.LocalDebug]: { result: "success", time: new Date(0) },
            },
            codeModifiedTime: {
              source: new Date(),
            },
          },
        } as WholeStatus)
      );
    });

    it("local debug succeeded after source changed", () => {
      chai.assert.isTrue(
        condition.isDebugSucceededAfterSourceCodeChanged({
          projectOpened: {
            actionStatus: {
              [CommandKey.LocalDebug]: { result: "success", time: new Date() },
            },
            codeModifiedTime: {
              source: new Date(0),
            },
          },
        } as WholeStatus)
      );
    });
  });

  describe("canPreviewInTestTool", () => {
    it("no opened project", () => {
      chai.assert.isFalse(condition.canPreviewInTestTool({} as WholeStatus));
    });

    it("no launch.json file", () => {
      chai.assert.isFalse(
        condition.canPreviewInTestTool({
          projectOpened: {},
        } as WholeStatus)
      );
    });

    it("no 'Test Tool' in launch.json file", () => {
      chai.assert.isFalse(
        condition.canPreviewInTestTool({
          projectOpened: {
            launchJSONContent: "123123123",
          },
        } as WholeStatus)
      );
    });

    it("'Test Tool' in launch.json file", () => {
      chai.assert.isTrue(
        condition.canPreviewInTestTool({
          projectOpened: {
            launchJSONContent: "Test Tool",
          },
        } as WholeStatus)
      );
    });
  });

  it("isM365AccountLogin", () => {
    chai.assert.isTrue(
      condition.isM365AccountLogin({
        machineStatus: {
          m365LoggedIn: true,
        },
      } as WholeStatus)
    );
    chai.assert.isFalse(
      condition.isM365AccountLogin({
        machineStatus: {
          m365LoggedIn: false,
        },
      } as WholeStatus)
    );
  });

  describe("isProvisionedSucceeded AfterInfraCodeChanged", () => {
    it("no opened project", () => {
      chai.assert.isFalse(condition.isProvisionedSucceededAfterInfraCodeChanged({} as WholeStatus));
    });

    it("provision not run before", () => {
      chai.assert.isFalse(
        condition.isProvisionedSucceededAfterInfraCodeChanged({
          projectOpened: {
            actionStatus: {
              [CommandKey.Provision]: { result: "no run", time: new Date() },
            },
          },
        } as WholeStatus)
      );
    });

    it("provision failed before", () => {
      chai.assert.isFalse(
        condition.isProvisionedSucceededAfterInfraCodeChanged({
          projectOpened: {
            actionStatus: {
              [CommandKey.Provision]: { result: "fail", time: new Date() },
            },
          },
        } as WholeStatus)
      );
    });

    it("provision succeeded before but out of date", () => {
      chai.assert.isFalse(
        condition.isProvisionedSucceededAfterInfraCodeChanged({
          projectOpened: {
            actionStatus: {
              [CommandKey.Provision]: { result: "success", time: new Date(0) },
            },
            codeModifiedTime: {
              infra: new Date(),
            },
          },
        } as WholeStatus)
      );
    });

    it("provision succeeded after infra changed", () => {
      chai.assert.isTrue(
        condition.isProvisionedSucceededAfterInfraCodeChanged({
          projectOpened: {
            actionStatus: {
              [CommandKey.Provision]: { result: "success", time: new Date() },
            },
            codeModifiedTime: {
              infra: new Date(0),
            },
          },
        } as WholeStatus)
      );
    });
  });

  it("isAzureAccountLogin", () => {
    chai.assert.isTrue(
      condition.isAzureAccountLogin({
        machineStatus: {
          azureLoggedIn: true,
        },
      } as WholeStatus)
    );
    chai.assert.isFalse(
      condition.isAzureAccountLogin({
        machineStatus: {
          azureLoggedIn: false,
        },
      } as WholeStatus)
    );
  });

  describe("isDeployed AfterSourceCodeChanged", () => {
    it("no opened project", () => {
      chai.assert.isFalse(condition.isDeployedAfterSourceCodeChanged({} as WholeStatus));
    });

    it("deploy not run before", () => {
      chai.assert.isFalse(
        condition.isDeployedAfterSourceCodeChanged({
          projectOpened: {
            actionStatus: {
              [CommandKey.Deploy]: { result: "no run", time: new Date() },
            },
          },
        } as WholeStatus)
      );
    });

    it("deploy failed before", () => {
      chai.assert.isFalse(
        condition.isDeployedAfterSourceCodeChanged({
          projectOpened: {
            actionStatus: {
              [CommandKey.Deploy]: { result: "fail", time: new Date() },
            },
          },
        } as WholeStatus)
      );
    });

    it("deploy succeeded before but out of date", () => {
      chai.assert.isFalse(
        condition.isDeployedAfterSourceCodeChanged({
          projectOpened: {
            actionStatus: {
              [CommandKey.Deploy]: { result: "success", time: new Date(0) },
            },
            codeModifiedTime: {
              source: new Date(),
            },
          },
        } as WholeStatus)
      );
    });

    it("deploy succeeded after source changed", () => {
      chai.assert.isTrue(
        condition.isDeployedAfterSourceCodeChanged({
          projectOpened: {
            actionStatus: {
              [CommandKey.Deploy]: { result: "success", time: new Date() },
            },
            codeModifiedTime: {
              source: new Date(0),
            },
          },
        } as WholeStatus)
      );
    });
  });

  describe("isPublishedSucceededBefore", () => {
    it("no opened project", () => {
      chai.assert.isFalse(condition.isPublishedSucceededBefore({} as WholeStatus));
    });

    it("publish not run before", () => {
      chai.assert.isFalse(
        condition.isPublishedSucceededBefore({
          projectOpened: {
            actionStatus: {
              [CommandKey.Publish]: { result: "no run", time: new Date() },
            },
          },
        } as WholeStatus)
      );
    });

    it("publish failed before", () => {
      chai.assert.isFalse(
        condition.isPublishedSucceededBefore({
          projectOpened: {
            actionStatus: {
              [CommandKey.Publish]: { result: "fail", time: new Date() },
            },
          },
        } as WholeStatus)
      );
    });

    it("publish succeeded", () => {
      chai.assert.isTrue(
        condition.isPublishedSucceededBefore({
          projectOpened: {
            actionStatus: {
              [CommandKey.Publish]: { result: "success", time: new Date() },
            },
          },
        } as WholeStatus)
      );
    });
  });

  describe("isHaveReadMe", () => {
    it("no opened project", () => {
      chai.assert.isFalse(condition.isHaveReadMe({} as WholeStatus));
    });

    it("no readme", () => {
      chai.assert.isFalse(
        condition.isHaveReadMe({
          projectOpened: {},
        } as WholeStatus)
      );
    });

    it("had readme", () => {
      chai.assert.isTrue(
        condition.isHaveReadMe({
          projectOpened: {
            readmeContent: "123123",
          },
        } as WholeStatus)
      );
    });
  });
});
