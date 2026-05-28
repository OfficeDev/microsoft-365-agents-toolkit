import * as sinon from "sinon";
import * as chai from "chai";
import { LocalEnvManager } from "@microsoft/teamsfx-core";
import { getNpmInstallLogInfo, getTestToolLogInfo } from "../../src/utils/localEnvManagerUtils";
import * as globalVariables from "../../src/globalVariables";

describe("LocalEnvUtils", () => {
  const sandbox = sinon.createSandbox();

  afterEach(() => {
    sandbox.restore();
  });

  it("Get NPM Install Log Info", async () => {
    const fakeNpmInstallLogInfo = {
      logFile: "NPM Install Log File",
      timestamp: new Date(),
      nodeVersion: undefined,
      npmVersion: undefined,
      cwd: undefined,
      exitCode: undefined,
      errorMessage: undefined,
    };
    sandbox.stub(LocalEnvManager.prototype, "getNpmInstallLogInfo").resolves(fakeNpmInstallLogInfo);
    const result = await getNpmInstallLogInfo();
    chai.expect(result).to.deep.equal(fakeNpmInstallLogInfo);
  });

  it("Get Test Tool Log Info", async () => {
    const fakeTestToolLogInfo = "Test Tool Log Info";
    sandbox.stub(globalVariables, "workspaceUri").value({ fsPath: "fakePath" });
    sandbox.stub(LocalEnvManager.prototype, "getTestToolLogInfo").resolves(fakeTestToolLogInfo);
    const result = await getTestToolLogInfo();
    chai.expect(result).to.equal(fakeTestToolLogInfo);
  });

  it("Get Test Tool Log Info and Return Undefined", async () => {
    const fakeTestToolLogInfo = "Test Tool Log Info";
    sandbox.stub(globalVariables, "workspaceUri");
    sandbox.stub(LocalEnvManager.prototype, "getTestToolLogInfo").resolves(fakeTestToolLogInfo);
    const result = await getTestToolLogInfo();
    chai.expect(result).to.be.undefined;
  });
});
