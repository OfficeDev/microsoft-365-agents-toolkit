import * as sinon from "sinon";
import * as chai from "chai";
import * as vscode from "vscode";
import * as globalVariables from "../../src/globalVariables";
import * as vsc_ui from "../../src/qm/vsc_ui";
import { MockCore } from "../mocks/mockCore";
import { ok, err, UserError } from "@microsoft/teamsfx-api";
import { CollaborationState } from "@microsoft/teamsfx-core";
import VsCodeLogInstance from "../../src/commonlib/log";
import { manageCollaboratorHandler } from "../../src/handlers/collaboratorHandlers";
import { ExtTelemetry } from "../../src/telemetry/extTelemetry";

describe("manageCollaboratorHandler", () => {
  const sandbox = sinon.createSandbox();

  beforeEach(() => {
    sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
    sandbox.stub(ExtTelemetry, "sendTelemetryErrorEvent");
    sandbox.stub(VsCodeLogInstance, "outputChannel").value({
      name: "name",
      append: (value: string) => {},
      appendLine: (value: string) => {},
      replace: (value: string) => {},
      clear: () => {},
      show: (...params: any[]) => {},
      hide: () => {},
      dispose: () => {},
    });
  });

  afterEach(() => {
    sandbox.restore();
  });

  it("happy path: grant permission", async () => {
    sandbox.stub(globalVariables, "core").value(new MockCore());
    sandbox.stub(vsc_ui, "VS_CODE_UI").value({
      selectOption: () => Promise.resolve(ok({ type: "success", result: "grantPermission" })),
    });
    sandbox.stub(MockCore.prototype, "grantPermission").returns(
      Promise.resolve(
        ok({
          state: CollaborationState.OK,
          userInfo: {
            userObjectId: "fake-user-object-id",
            userPrincipalName: "fake-user-principle-name",
          },
          permissions: [
            {
              name: "name",
              type: "type",
              resourceId: "id",
              roles: ["Owner"],
            },
          ],
        })
      )
    );

    const result = await manageCollaboratorHandler("env");
    chai.expect(result.isOk()).equals(true);
  });

  it("happy path: list collaborator", async () => {
    sandbox.stub(globalVariables, "core").value(new MockCore());
    sandbox.stub(vsc_ui, "VS_CODE_UI").value({
      selectOption: () => Promise.resolve(ok({ type: "success", result: "listCollaborator" })),
    });
    sandbox.stub(MockCore.prototype, "listCollaborator").returns(
      Promise.resolve(
        ok({
          state: CollaborationState.OK,
          collaborators: [
            {
              userPrincipalName: "userPrincipalName",
              userObjectId: "userObjectId",
              isAadOwner: true,
              teamsAppResourceId: "teamsAppResourceId",
            },
          ],
        })
      )
    );

    const result = await manageCollaboratorHandler("env");
    chai.expect(result.isOk()).equals(true);
  });

  it("happy path: list collaborator throws error", async () => {
    sandbox.stub(globalVariables, "core").value(new MockCore());
    sandbox.stub(vsc_ui, "VS_CODE_UI").value({
      selectOption: () => Promise.resolve(ok({ type: "success", result: "listCollaborator" })),
    });
    sandbox.stub(MockCore.prototype, "listCollaborator").throws(new Error("Error"));

    const result = await manageCollaboratorHandler("env");
    chai.expect(result.isErr()).equals(true);
  });

  it("happy path: list collaborator throws login error", async () => {
    sandbox.stub(globalVariables, "core").value(new MockCore());
    sandbox.stub(vsc_ui, "VS_CODE_UI").value({
      selectOption: () => Promise.resolve(ok({ type: "success", result: "listCollaborator" })),
    });
    const showErrorMessageStub = sandbox
      .stub(vscode.window, "showErrorMessage")
      .resolves(undefined);
    sandbox
      .stub(globalVariables.core, "listCollaborator")
      .throws(new Error("Cannot get user login information"));

    const result = await manageCollaboratorHandler("env");
    chai.expect(result.isErr()).equals(true);
    chai.assert.isTrue(showErrorMessageStub.called);
  });

  it("User Cancel", async () => {
    sandbox.stub(globalVariables, "core").value(new MockCore());
    sandbox.stub(vsc_ui, "VS_CODE_UI").value({
      selectOption: () =>
        Promise.resolve(err(new UserError("source", "errorName", "errorMessage"))),
    });

    const result = await manageCollaboratorHandler();
    chai.expect(result.isErr()).equals(true);
  });
});
