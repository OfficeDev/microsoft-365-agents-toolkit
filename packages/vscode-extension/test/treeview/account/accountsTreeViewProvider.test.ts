import * as chai from "chai";
import * as sinon from "sinon";
import { stubInterface } from "ts-sinon";

import { AzureAccountProvider, M365TokenProvider, ok, TokenRequest } from "@microsoft/teamsfx-api";
import * as projectSettingsHelper from "@microsoft/teamsfx-core/build/common/projectSettingsHelper";

import * as globalVariables from "../../../src/globalVariables";
import AccountTreeViewProvider from "../../../src/treeview/account/accountTreeViewProvider";
import EnvironemtTreeProvider from "../../../src/treeview/environmentTreeViewProvider";

describe("AccountTreeViewProvider", () => {
  const sandbox = sinon.createSandbox();

  afterEach(() => {
    sandbox.restore();
  });

  it("subscribeToStatusChanges", async () => {
    sandbox.stub(projectSettingsHelper, "isValidProject").returns(true);
    sandbox.stub(globalVariables, "workspaceUri").value({ fsPath: "test" });
    sandbox.stub(EnvironemtTreeProvider, "reloadEnvironments");
    const azureAccountProviderStub = stubInterface<AzureAccountProvider>();
    const m365TokenProviderStub = stubInterface<M365TokenProvider>();

    let m365StatusChange: (
      status: string,
      token?: string,
      accountInfo?: Record<string, unknown>
    ) => Promise<void> = () => Promise.resolve();
    m365TokenProviderStub.setStatusChangeMap.callsFake(
      (
        name: string,
        tokenRequest: TokenRequest,
        statusChange: (
          status: string,
          token?: string,
          accountInfo?: Record<string, unknown>
        ) => Promise<void>
      ) => {
        m365StatusChange = statusChange;
        return Promise.resolve(ok(true));
      }
    );
    let azureStatusChange: (
      status: string,
      token?: string,
      accountInfo?: Record<string, unknown>
    ) => Promise<void> = () => Promise.resolve();
    azureAccountProviderStub.setStatusChangeMap.callsFake(
      (
        name: string,
        statusChange: (
          status: string,
          token?: string,
          accountInfo?: Record<string, unknown>
        ) => Promise<void>
      ) => {
        azureStatusChange = statusChange;
        return Promise.resolve(true);
      }
    );

    AccountTreeViewProvider.subscribeToStatusChanges({
      azureAccountProvider: azureAccountProviderStub,
      m365TokenProvider: m365TokenProviderStub,
    });

    chai.assert.isTrue(azureAccountProviderStub.setStatusChangeMap.calledOnce);
    chai.assert.isTrue(m365TokenProviderStub.setStatusChangeMap.calledOnce);

    const m365SigingInStub = sandbox.stub(AccountTreeViewProvider.m365AccountNode, "setSigningIn");
    await m365StatusChange("SigningIn");
    chai.assert.isTrue(m365SigingInStub.calledOnce);

    const m365SignedOutStub = sandbox.stub(AccountTreeViewProvider.m365AccountNode, "setSignedOut");
    await m365StatusChange("SignedOut");
    chai.assert.isTrue(m365SignedOutStub.calledOnce);

    const m365SignedInStub = sandbox.stub(AccountTreeViewProvider.m365AccountNode, "setSignedIn");
    const updateChecksStub = sandbox.stub(AccountTreeViewProvider.m365AccountNode, "updateChecks");
    await m365StatusChange("SignedIn", "token", { upn: "upn" });
    chai.assert.isTrue(m365SignedInStub.calledOnceWithExactly("upn", ""));
    chai.assert.isTrue(updateChecksStub.calledOnce);

    m365SignedInStub.reset();
    updateChecksStub.reset();
    await m365StatusChange("SignedIn", "token", { tid: "tid" });
    chai.assert.isTrue(m365SignedInStub.calledOnceWithExactly("", "tid"));
    chai.assert.isTrue(updateChecksStub.calledOnce);

    m365SignedInStub.reset();
    updateChecksStub.reset();
    await m365StatusChange("SignedIn", "token", { upn: "upn", tid: "tid" });
    chai.assert.isTrue(m365SignedInStub.calledOnceWithExactly("upn", "tid"));
    chai.assert.isTrue(updateChecksStub.calledOnce);

    const m365SwitchingStub = sandbox.stub(AccountTreeViewProvider.m365AccountNode, "setSwitching");
    await m365StatusChange("Switching");
    chai.assert.isTrue(m365SwitchingStub.calledOnce);

    const azureSignedOutStub = sandbox.stub(
      AccountTreeViewProvider.azureAccountNode,
      "setSignedOut"
    );
    await azureStatusChange("SignedOut");
    chai.assert.isTrue(azureSignedOutStub.calledOnce);

    const azureSignedInStub = sandbox.stub(AccountTreeViewProvider.azureAccountNode, "setSignedIn");
    await azureStatusChange("SignedIn", "token", { upn: "upn" });
    chai.assert.isTrue(azureSignedInStub.calledOnce);

    azureSignedInStub.reset();
    await azureStatusChange("SignedIn", "token", { upn: "upn", tid: "tid" });
    chai.assert.isTrue(azureSignedInStub.calledOnceWithExactly("token", "tid", "upn"));

    const azureSigningInStub = sandbox.stub(
      AccountTreeViewProvider.azureAccountNode,
      "setSigningIn"
    );
    await azureStatusChange("SigningIn", undefined, {});
    chai.assert.isTrue(azureSigningInStub.calledOnce);
  });

  it("getChildren", async () => {
    sandbox.stub(globalVariables, "isSPFxProject").value(false);

    const children = await AccountTreeViewProvider.getChildren();

    chai.assert.equal(children?.length, 2);
  });
});
