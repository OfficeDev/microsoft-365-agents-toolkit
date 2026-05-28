import { err, ok } from "@microsoft/teamsfx-api";
import { MosServiceScope, AppStudioScopes, PackageService } from "@microsoft/teamsfx-core";
import * as sinon from "sinon";
import * as vscode from "vscode";
import VsCodeLogInstance from "../../../src/commonlib/log";
import M365TokenInstance from "../../../src/commonlib/m365Login";
import { checkCopilotAccessHandler } from "../../../src/handlers/accounts/checkCopilotAccess";
import { MockLogProvider } from "../../mocks/mockTools";

describe("check copilot access", () => {
  const sandbox = sinon.createSandbox();

  beforeEach(() => {
    sandbox.stub(PackageService, "GetSharedInstance").returns(new PackageService("endpoint"));
  });

  afterEach(() => {
    sandbox.restore();
  });

  it("check copilot access in walkthrough: not signed in && with access", async () => {
    const copilotCheckServiceScope = process.env.SIDELOADING_SERVICE_SCOPE ?? MosServiceScope;
    const m365GetStatusStub = sandbox
      .stub(M365TokenInstance, "getStatus")
      .withArgs({ scopes: AppStudioScopes })
      .resolves(err({ error: "unknown" } as any));
    const m365GetAccessTokenStub = sandbox
      .stub(M365TokenInstance, "getAccessToken")
      .withArgs({ scopes: [copilotCheckServiceScope] })
      .resolves(ok("stubedString"));
    const getCopilotStatusStub = sandbox
      .stub(PackageService.prototype, "getCopilotStatus")
      .resolves(true);
    const showMessageStub = sandbox.stub(vscode.window, "showInformationMessage").resolves({
      title: "Sign in",
    } as vscode.MessageItem);
    const signInM365Stub = sandbox.stub(vscode.commands, "executeCommand").resolves();
    const semLogStub = sandbox.stub(VsCodeLogInstance, "semLog").resolves();

    await checkCopilotAccessHandler();

    sandbox.assert.calledOnce(m365GetStatusStub);
    sandbox.assert.calledOnce(showMessageStub);
    sandbox.assert.calledOnce(signInM365Stub);
    sandbox.assert.calledOnce(m365GetAccessTokenStub);
    sandbox.assert.calledOnce(getCopilotStatusStub);
    sandbox.assert.calledOnce(semLogStub);
  });

  it("check copilot access in walkthrough: not signed in && no access", async () => {
    const copilotCheckServiceScope = process.env.SIDELOADING_SERVICE_SCOPE ?? MosServiceScope;
    const m365GetStatusStub = sandbox
      .stub(M365TokenInstance, "getStatus")
      .withArgs({ scopes: AppStudioScopes })
      .resolves(err({ error: "unknown" } as any));
    const m365GetAccessTokenStub = sandbox
      .stub(M365TokenInstance, "getAccessToken")
      .withArgs({ scopes: [copilotCheckServiceScope] })
      .resolves(ok("stubedString"));

    const getCopilotStatusStub = sandbox
      .stub(PackageService.prototype, "getCopilotStatus")
      .resolves(false);

    const showMessageStub = sandbox.stub(vscode.window, "showInformationMessage").resolves({
      title: "Sign in",
    } as vscode.MessageItem);

    const signInM365Stub = sandbox.stub(vscode.commands, "executeCommand").resolves();

    const semLogStub = sandbox.stub(VsCodeLogInstance, "semLog").resolves();

    await checkCopilotAccessHandler();

    sandbox.assert.calledOnce(m365GetStatusStub);
    sandbox.assert.calledOnce(showMessageStub);
    sandbox.assert.calledOnce(signInM365Stub);
    sandbox.assert.calledOnce(m365GetAccessTokenStub);
    sandbox.assert.calledOnce(getCopilotStatusStub);
    sandbox.assert.calledOnce(semLogStub);
  });

  it("check copilot access in walkthrough: not signed in && throw error", async () => {
    const copilotCheckServiceScope = process.env.SIDELOADING_SERVICE_SCOPE ?? MosServiceScope;
    const m365GetStatusStub = sandbox
      .stub(M365TokenInstance, "getStatus")
      .withArgs({ scopes: AppStudioScopes })
      .resolves(err({ error: "unknown" } as any));
    sandbox
      .stub(M365TokenInstance, "getAccessToken")
      .withArgs({ scopes: [copilotCheckServiceScope] })
      .resolves(ok("stubedString"));

    sandbox.stub(PackageService.prototype, "getCopilotStatus").resolves(true);

    const showMessageStub = sandbox.stub(vscode.window, "showInformationMessage").resolves({
      title: "Sign in",
    } as vscode.MessageItem);

    const signInM365Stub = sandbox.stub(vscode.commands, "executeCommand").rejects(Error("error"));

    const result = await checkCopilotAccessHandler();

    sandbox.assert.calledOnce(m365GetStatusStub);
    sandbox.assert.calledOnce(showMessageStub);
    sandbox.assert.calledOnce(signInM365Stub);
    sandbox.assert.match(result.isErr() ? result.error.message : "", "error");
  });

  it("check copilot access in walkthrough: signed in && no access", async () => {
    const copilotCheckServiceScope = process.env.SIDELOADING_SERVICE_SCOPE ?? MosServiceScope;
    const m365GetStatusStub = sandbox
      .stub(M365TokenInstance, "getStatus")
      .withArgs({ scopes: AppStudioScopes })
      .resolves(ok({ status: "SignedIn", accountInfo: { upn: "test.email.com" } }));
    const m365GetAccessTokenStub = sandbox
      .stub(M365TokenInstance, "getAccessToken")
      .withArgs({ scopes: [copilotCheckServiceScope] })
      .resolves(ok("stubedString"));

    const getCopilotStatusStub = sandbox
      .stub(PackageService.prototype, "getCopilotStatus")
      .resolves(false);

    const showMessageStub = sandbox.stub(vscode.window, "showInformationMessage").resolves({
      title: "Sign in",
    } as vscode.MessageItem);

    const signInM365Stub = sandbox.stub(vscode.commands, "executeCommand").resolves();

    const semLogStub = sandbox.stub(VsCodeLogInstance, "semLog").resolves();

    await checkCopilotAccessHandler();

    sandbox.assert.calledOnce(m365GetStatusStub);
    sandbox.assert.notCalled(showMessageStub);
    sandbox.assert.notCalled(signInM365Stub);
    sandbox.assert.calledOnce(m365GetAccessTokenStub);
    sandbox.assert.calledOnce(getCopilotStatusStub);
    sandbox.assert.calledOnce(semLogStub);
  });

  it("check copilot access in walkthrough: signed in && with access", async () => {
    const copilotCheckServiceScope = process.env.SIDELOADING_SERVICE_SCOPE ?? MosServiceScope;
    const m365GetStatusStub = sandbox
      .stub(M365TokenInstance, "getStatus")
      .withArgs({ scopes: AppStudioScopes })
      .resolves(ok({ status: "SignedIn", accountInfo: { upn: "test.email.com" } }));
    const m365GetAccessTokenStub = sandbox
      .stub(M365TokenInstance, "getAccessToken")
      .withArgs({ scopes: [copilotCheckServiceScope] })
      .resolves(ok("stubedString"));

    const getCopilotStatusStub = sandbox
      .stub(PackageService.prototype, "getCopilotStatus")
      .resolves(true);

    const showMessageStub = sandbox.stub(vscode.window, "showInformationMessage").resolves({
      title: "Sign in",
    } as vscode.MessageItem);

    const signInM365Stub = sandbox.stub(vscode.commands, "executeCommand").resolves();

    const semLogStub = sandbox.stub(VsCodeLogInstance, "semLog").resolves();

    await checkCopilotAccessHandler();

    sandbox.assert.calledOnce(m365GetStatusStub);
    sandbox.assert.notCalled(showMessageStub);
    sandbox.assert.notCalled(signInM365Stub);
    sandbox.assert.calledOnce(m365GetAccessTokenStub);
    sandbox.assert.calledOnce(getCopilotStatusStub);
    sandbox.assert.calledOnce(semLogStub);
  });

  it("check copilot access in walkthrough: signed in && throw error", async () => {
    const copilotCheckServiceScope = process.env.SIDELOADING_SERVICE_SCOPE ?? MosServiceScope;
    const m365GetStatusStub = sandbox
      .stub(M365TokenInstance, "getStatus")
      .withArgs({ scopes: AppStudioScopes })
      .resolves(ok({ status: "SignedIn", accountInfo: { upn: "test.email.com" } }));
    const m365GetAccessTokenStub = sandbox
      .stub(M365TokenInstance, "getAccessToken")
      .withArgs({ scopes: [copilotCheckServiceScope] })
      .resolves(err({ error: "error" } as any));

    const result = await checkCopilotAccessHandler();

    sandbox.assert.calledOnce(m365GetStatusStub);
    sandbox.assert.calledOnce(m365GetAccessTokenStub);
    sandbox.assert.match(result.isErr() ? result.error : {}, { error: "error" });
  });
});
