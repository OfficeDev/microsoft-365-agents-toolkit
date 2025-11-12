import * as chai from "chai";
import * as sinon from "sinon";

import { ok } from "@microsoft/teamsfx-api";
import { environmentManager } from "@microsoft/teamsfx-core";
import * as projectSettingsHelper from "@microsoft/teamsfx-core/build/src/common/projectSettingsHelper";

import * as globalVariables from "../../src/globalVariables";
import EnvironmentTreeViewProvider from "../../src/treeview/environmentTreeViewProvider";

describe("EnvironmentTreeViewProvider", () => {
  const sandbox = sinon.createSandbox();

  afterEach(() => {
    sandbox.restore();
  });

  it("reloadEnvironments", async () => {
    sandbox.stub(projectSettingsHelper, "isValidProject").returns(true);
    sandbox.stub(globalVariables, "workspaceUri").value({ fsPath: "test" });

    const result = await EnvironmentTreeViewProvider.reloadEnvironments();

    chai.assert.isTrue(result.isOk());
  });

  it("getChildren", async () => {
    sandbox.stub(globalVariables, "workspaceUri").value({ fsPath: "test" });
    sandbox.stub(environmentManager, "listRemoteEnvConfigs").returns(Promise.resolve(ok(["test"])));
    sandbox.stub(environmentManager, "getExistingNonRemoteEnvs").returns(Promise.resolve(["test"]));

    const children = await EnvironmentTreeViewProvider.getChildren();

    chai.assert.equal(children?.length, 2);
  });
});
