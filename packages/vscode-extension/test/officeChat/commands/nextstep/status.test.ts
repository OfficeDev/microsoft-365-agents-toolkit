import * as chai from "chai";
import chaiPromised from "chai-as-promised";
import * as sinon from "sinon";
import * as helper from "../../../../src/chat/commands/nextstep/helper";
import * as projectStatusUtils from "../../../../src/utils/projectStatusUtils";
import * as status from "../../../../src/officeChat/commands/nextStep/status";
import fs from "fs-extra";
import child_process from "child_process";
import { OfficeWholeStatus } from "../../../../src/officeChat/commands/nextStep/types";

chai.use(chaiPromised);

describe("office steps: getWholeStatus", () => {
  const sandbox = sinon.createSandbox();
  afterEach(() => {
    sandbox.restore();
  });

  it("folder !== undefined and node exists", async () => {
    sandbox.stub(helper, "getProjectMetadata").returns({ projectId: "test-id" });
    sandbox
      .stub(projectStatusUtils, "getProjectStatus")
      .resolves(projectStatusUtils.emptyProjectStatus());
    sandbox.stub(projectStatusUtils, "getFileModifiedTime").resolves(new Date(0));
    sandbox.stub(projectStatusUtils, "getREADME").resolves(undefined);
    sandbox.stub(projectStatusUtils, "getLaunchJSON").resolves(undefined);
    sandbox.stub(helper, "checkCredential").resolves({ m365LoggedIn: true, azureLoggedIn: true });
    sandbox.stub(helper, "globalStateGet").resolves(true);
    sandbox.stub(helper, "globalStateUpdate");
    sandbox.stub(fs, "pathExists").resolves(true);
    sandbox.stub(child_process, "exec").yields(null);
    await chai.expect(status.getWholeStatus("test-folder")).to.eventually.deep.equal({
      machineStatus: {
        azureLoggedIn: true,
        firstInstalled: true,
        m365LoggedIn: true,
      },
      projectOpened: {
        path: "test-folder",
        projectId: "test-id",
        codeModifiedTime: {
          source: new Date(0),
          infra: new Date(0),
        },
        actionStatus: projectStatusUtils.emptyProjectStatus(),
        readmeContent: undefined,
        launchJSONContent: undefined,
        nodeModulesExist: true,
        isNodeInstalled: true,
      },
    } as OfficeWholeStatus);
  });

  it("folder !== undefined and node not exists", async () => {
    sandbox.stub(helper, "getProjectMetadata").returns({ projectId: "test-id" });
    sandbox
      .stub(projectStatusUtils, "getProjectStatus")
      .resolves(projectStatusUtils.emptyProjectStatus());
    sandbox.stub(projectStatusUtils, "getFileModifiedTime").resolves(new Date(0));
    sandbox.stub(projectStatusUtils, "getREADME").resolves(undefined);
    sandbox.stub(projectStatusUtils, "getLaunchJSON").resolves(undefined);
    sandbox.stub(helper, "checkCredential").resolves({ m365LoggedIn: true, azureLoggedIn: true });
    sandbox.stub(helper, "globalStateGet").resolves(true);
    sandbox.stub(helper, "globalStateUpdate");
    sandbox.stub(fs, "pathExists").resolves(true);
    sandbox.stub(child_process, "exec").yields(new Error("error"), null, null);
    await chai.expect(status.getWholeStatus("test-folder")).to.eventually.deep.equal({
      machineStatus: {
        azureLoggedIn: true,
        firstInstalled: true,
        m365LoggedIn: true,
      },
      projectOpened: {
        path: "test-folder",
        projectId: "test-id",
        codeModifiedTime: {
          source: new Date(0),
          infra: new Date(0),
        },
        actionStatus: projectStatusUtils.emptyProjectStatus(),
        readmeContent: undefined,
        launchJSONContent: undefined,
        nodeModulesExist: true,
        isNodeInstalled: false,
      },
    } as OfficeWholeStatus);
  });
});
