import * as chai from "chai";
import { vi } from "vitest";
import { createMock } from "../../mocks/vitestMockUtils";
import { mockValue } from "../../mocks/vitestMockUtils";

import { AzureAccountProvider, M365TokenProvider, ok, TokenRequest } from "@microsoft/teamsfx-api";
import { FeatureFlags, GraphScopes, featureFlagManager } from "@microsoft/teamsfx-core";
import * as projectSettingsHelper from "@microsoft/teamsfx-core/build/common/projectSettingsHelper";

import * as globalVariables from "../../../src/globalVariables";
import AccountTreeViewProvider from "../../../src/treeview/account/accountTreeViewProvider";
import EnvironemtTreeProvider from "../../../src/treeview/environmentTreeViewProvider";

describe("AccountTreeViewProvider", () => {
  it("subscribeToStatusChanges", async () => {
    vi.spyOn(projectSettingsHelper, "isValidProject").mockReturnValue(true);
    mockValue(globalVariables, "workspaceUri", { fsPath: "test" });
    vi.spyOn(EnvironemtTreeProvider, "reloadEnvironments");
    const azureAccountProviderStub = createMock<AzureAccountProvider>();
    const m365TokenProviderStub = createMock<M365TokenProvider>();

    let m365StatusChange: (
      status: string,
      token?: string,
      accountInfo?: Record<string, unknown>
    ) => Promise<void> = () => Promise.resolve();
    m365TokenProviderStub.setStatusChangeMap.mockImplementation(
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
    azureAccountProviderStub.setStatusChangeMap.mockImplementation(
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

    const m365SigingInStub = vi.spyOn(AccountTreeViewProvider.m365AccountNode, "setSigningIn");
    await m365StatusChange("SigningIn");
    chai.assert.isTrue(m365SigingInStub.calledOnce);

    const m365SignedOutStub = vi.spyOn(AccountTreeViewProvider.m365AccountNode, "setSignedOut");
    await m365StatusChange("SignedOut");
    chai.assert.isTrue(m365SignedOutStub.calledOnce);

    const m365SignedInStub = vi
      .spyOn(AccountTreeViewProvider.m365AccountNode, "setSignedIn")
      .mockResolvedValue();
    const updateChecksStub = vi.spyOn(AccountTreeViewProvider.m365AccountNode, "updateChecks");
    await m365StatusChange("SignedIn", "token", { upn: "upn" });
    chai.assert.isTrue(m365SignedInStub.calledOnceWithExactly("upn", ""));
    chai.assert.isTrue(updateChecksStub.calledOnce);

    m365SignedInStub.mockClear();
    updateChecksStub.mockClear();
    await m365StatusChange("SignedIn", "token", { tid: "tid" });
    chai.assert.isTrue(m365SignedInStub.calledOnceWithExactly("", "tid"));
    chai.assert.isTrue(updateChecksStub.calledOnce);

    m365SignedInStub.mockClear();
    updateChecksStub.mockClear();
    await m365StatusChange("SignedIn", "token", { upn: "upn", tid: "tid" });
    chai.assert.isTrue(m365SignedInStub.calledOnceWithExactly("upn", "tid"));
    chai.assert.isTrue(updateChecksStub.calledOnce);

    const m365SwitchingStub = vi.spyOn(AccountTreeViewProvider.m365AccountNode, "setSwitching");
    await m365StatusChange("Switching");
    chai.assert.isTrue(m365SwitchingStub.calledOnce);

    const azureSignedOutStub = vi.spyOn(AccountTreeViewProvider.azureAccountNode, "setSignedOut");
    await azureStatusChange("SignedOut");
    chai.assert.isTrue(azureSignedOutStub.calledOnce);

    const azureSignedInStub = vi
      .spyOn(AccountTreeViewProvider.azureAccountNode, "setSignedIn")
      .mockResolvedValue();
    await azureStatusChange("SignedIn", "token", { upn: "upn" });
    chai.assert.isTrue(azureSignedInStub.calledOnce);

    azureSignedInStub.mockClear();
    await azureStatusChange("SignedIn", "token", { upn: "upn", tid: "tid" });
    chai.assert.isTrue(azureSignedInStub.calledOnceWithExactly("token", "tid", "upn"));

    const azureSigningInStub = vi.spyOn(AccountTreeViewProvider.azureAccountNode, "setSigningIn");
    await azureStatusChange("SigningIn", undefined, {});
    chai.assert.isTrue(azureSigningInStub.calledOnce);
  });

  it("getChildren", async () => {
    mockValue(globalVariables, "isSPFxProject", false);

    const children = await AccountTreeViewProvider.getChildren();

    chai.assert.equal(children?.length, 2);
  });

  it("subscribeToStatusChanges uses Graph scopes in sovereign high", async () => {
    vi.spyOn(featureFlagManager, "getStringValue").mockReturnValue("DoD");
    const azureAccountProviderStub = createMock<AzureAccountProvider>();
    const m365TokenProviderStub = createMock<M365TokenProvider>();

    m365TokenProviderStub.setStatusChangeMap.mockReturnValue(Promise.resolve(ok(true)));
    azureAccountProviderStub.setStatusChangeMap.mockReturnValue(Promise.resolve(true));

    AccountTreeViewProvider.subscribeToStatusChanges({
      azureAccountProvider: azureAccountProviderStub,
      m365TokenProvider: m365TokenProviderStub,
    });

    chai.assert.isTrue(
      m365TokenProviderStub.setStatusChangeMap.calledWithMatch(
        "tree-view",
        { scopes: GraphScopes },
        expect.any(Function)
      )
    );
  });
});
