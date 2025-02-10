import * as sinon from "sinon";
import * as chai from "chai";
import * as globalVariables from "../../src/globalVariables";
import * as telemetry from "../../src/telemetry/extTelemetry";
import * as workspaceUtils from "../../src/utils/workspaceUtils";
import { createDeclarativeAgentWithApiSpec } from "../../src/handlers/createDeclarativeAgentWithApiSpecHandler";
import { err, UserError } from "@microsoft/teamsfx-api";
import { MockCore } from "../mocks/mockCore";

describe("createDeclarativeAgentWithApiSpecHandler", () => {
  const sandbox = sinon.createSandbox();

  beforeEach(() => {
    sandbox.stub(telemetry.ExtTelemetry, "sendTelemetryEvent");
    sandbox.stub(telemetry.ExtTelemetry, "sendTelemetryErrorEvent");
  });

  afterEach(() => {
    sandbox.restore();
  });

  it("should return error if args are invalid", async () => {
    const core = new MockCore();
    sandbox.stub(globalVariables, "core").value(core);
    const openFolder = sandbox.stub(workspaceUtils, "openFolder").resolves();

    const res = await createDeclarativeAgentWithApiSpec([]);

    chai.assert.isTrue(res.isErr());
    chai.assert.isTrue(openFolder.notCalled);
    if (res.isErr()) {
      chai.assert.equal(res.error.name, "invalidParameter");
    }
  });

  it("should create project successfully with valid args", async () => {
    const core = new MockCore();
    sandbox.stub(globalVariables, "core").value(core);
    const openFolder = sandbox.stub(workspaceUtils, "openFolder").resolves();

    const res = await createDeclarativeAgentWithApiSpec(["test-path"]);

    chai.assert.isTrue(res.isOk());
    chai.assert.isTrue(openFolder.calledOnce);
  });

  it("should throw error if core return error", async () => {
    const core = new MockCore();
    sandbox.stub(globalVariables, "core").value(core);
    sandbox
      .stub(globalVariables.core, "createProject")
      .resolves(err(new UserError("core", "fakeError", "fakeErrorMessage")));
    const openFolder = sandbox.stub(workspaceUtils, "openFolder").resolves();

    const res = await createDeclarativeAgentWithApiSpec(["test-path"]);

    chai.assert.isTrue(res.isErr());
    chai.assert.isTrue(openFolder.notCalled);
    if (res.isErr()) {
      chai.assert.equal(res.error.name, "fakeError");
    }
  });
});
