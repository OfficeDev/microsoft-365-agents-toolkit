/**
 * Copyright (c) Microsoft Corporation.
 * Licensed under the MIT license.
 */

import { expect } from "chai";
import { describe, it, afterEach } from "mocha";
import * as sinon from "sinon";
import { ok, err } from "neverthrow";
import {
  ensureM365Auth,
  ensureAzureAuth,
  ensureSubscription,
  ensureResourceGroup,
  ensureResourceSuffix,
  confirmProvision,
  confirmDeploy,
} from "../../../src/lifecycle/prerequisites";
import type { AtkContext } from "../../../src/core/context";

// ---------------------------------------------------------------------------
// Mock factory
// ---------------------------------------------------------------------------

function createMockContext(overrides?: Partial<AtkContext>): AtkContext {
  return {
    auth: {
      m365TokenProvider: {
        getAccessToken: sinon.stub().resolves(ok("token")),
        getJsonObject: sinon.stub().resolves(ok({})),
        getStatus: sinon.stub().resolves(ok({ status: "SignedIn" })),
        signout: sinon.stub().resolves(true),
        switchTenant: sinon.stub(),
        setStatusChangeMap: sinon.stub(),
        removeStatusChangeMap: sinon.stub(),
      },
      azureAccountProvider: {
        getIdentityCredentialAsync: sinon.stub().resolves({}),
        signout: sinon.stub().resolves(true),
        switchTenant: sinon.stub(),
        setStatusChangeMap: sinon.stub(),
        removeStatusChangeMap: sinon.stub(),
        getJsonObject: sinon.stub().resolves(undefined),
        listSubscriptions: sinon.stub().resolves([]),
        setSubscription: sinon.stub().resolves(),
        getAccountInfo: sinon.stub().returns(undefined),
        getSelectedSubscription: sinon.stub().resolves(undefined),
      },
    },
    logger: {
      debug: sinon.stub(),
      info: sinon.stub(),
      warning: sinon.stub(),
      error: sinon.stub(),
    },
    telemetry: {
      sendTelemetryEvent: sinon.stub(),
      sendTelemetryErrorEvent: sinon.stub(),
      sendTelemetryException: sinon.stub(),
    },
    ui: {
      selectOption: sinon.stub(),
      selectOptions: sinon.stub(),
      inputText: sinon.stub(),
      selectFile: sinon.stub(),
      selectFiles: sinon.stub(),
      selectFolder: sinon.stub(),
      openUrl: sinon.stub(),
      showMessage: sinon.stub(),
      createProgressBar: sinon.stub(),
      confirm: sinon.stub(),
    },
    correlationId: "test-corr-id",
    ...overrides,
  } as any;
}

// ---------------------------------------------------------------------------
// ensureM365Auth
// ---------------------------------------------------------------------------

describe("ensureM365Auth", () => {
  afterEach(() => sinon.restore());

  it("should return tenant info on success", async () => {
    const ctx = createMockContext();
    (ctx.auth.m365TokenProvider.getJsonObject as sinon.SinonStub).resolves(
      ok({ tid: "tenant-1", name: "Alice", preferred_username: "alice@contoso.com" })
    );

    const result = await ensureM365Auth(ctx);

    expect(result.isOk()).to.be.true;
    if (result.isOk()) {
      expect(result.value.tenantId).to.equal("tenant-1");
      expect(result.value.displayName).to.equal("Alice");
    }
  });

  it("should use tenantId claim when tid is missing", async () => {
    const ctx = createMockContext();
    (ctx.auth.m365TokenProvider.getJsonObject as sinon.SinonStub).resolves(
      ok({ tenantId: "tenant-2", upn: "bob@contoso.com" })
    );

    const result = await ensureM365Auth(ctx);

    expect(result.isOk()).to.be.true;
    if (result.isOk()) {
      expect(result.value.tenantId).to.equal("tenant-2");
      expect(result.value.displayName).to.equal("bob@contoso.com");
    }
  });

  it("should return error when token acquisition fails", async () => {
    const ctx = createMockContext();
    (ctx.auth.m365TokenProvider.getJsonObject as sinon.SinonStub).resolves(
      err({ code: "LoginFailed", message: "User cancelled", kind: "user", source: "test" })
    );

    const result = await ensureM365Auth(ctx);

    expect(result.isErr()).to.be.true;
    if (result.isErr()) {
      expect(result.error.code).to.equal("M365AuthFailed");
    }
  });

  it("should return error when tenant ID is empty", async () => {
    const ctx = createMockContext();
    (ctx.auth.m365TokenProvider.getJsonObject as sinon.SinonStub).resolves(
      ok({ name: "NoTenant" })
    );

    const result = await ensureM365Auth(ctx);

    expect(result.isErr()).to.be.true;
    if (result.isErr()) {
      expect(result.error.code).to.equal("M365TenantNotFound");
    }
  });
});

// ---------------------------------------------------------------------------
// ensureAzureAuth
// ---------------------------------------------------------------------------

describe("ensureAzureAuth", () => {
  afterEach(() => sinon.restore());

  it("should return account info on success", async () => {
    const ctx = createMockContext();
    (ctx.auth.azureAccountProvider.getIdentityCredentialAsync as sinon.SinonStub).resolves({});
    (ctx.auth.azureAccountProvider.getAccountInfo as sinon.SinonStub).returns({
      email: "alice@contoso.com",
      tenantId: "az-tenant-1",
    });

    const result = await ensureAzureAuth(ctx);

    expect(result.isOk()).to.be.true;
    if (result.isOk()) {
      expect(result.value.accountId).to.equal("alice@contoso.com");
      expect(result.value.tenantId).to.equal("az-tenant-1");
    }
  });

  it("should return error when credential is null", async () => {
    const ctx = createMockContext();
    (ctx.auth.azureAccountProvider.getIdentityCredentialAsync as sinon.SinonStub).resolves(null);

    const result = await ensureAzureAuth(ctx);

    expect(result.isErr()).to.be.true;
    if (result.isErr()) {
      expect(result.error.code).to.equal("AzureAuthFailed");
    }
  });

  it("should handle undefined getAccountInfo", async () => {
    const ctx = createMockContext();
    (ctx.auth.azureAccountProvider.getIdentityCredentialAsync as sinon.SinonStub).resolves({});
    (ctx.auth.azureAccountProvider as any).getAccountInfo = undefined;

    const result = await ensureAzureAuth(ctx);

    expect(result.isOk()).to.be.true;
    if (result.isOk()) {
      expect(result.value.accountId).to.be.undefined;
    }
  });
});

// ---------------------------------------------------------------------------
// ensureSubscription
// ---------------------------------------------------------------------------

describe("ensureSubscription", () => {
  afterEach(() => sinon.restore());

  it("should return existing subscription from envMap", async () => {
    const ctx = createMockContext();
    const envMap = new Map([
      ["AZURE_SUBSCRIPTION_ID", "sub-123"],
      ["AZURE_SUBSCRIPTION_NAME", "My Sub"],
      ["AZURE_TENANT_ID", "tenant-456"],
    ]);

    const result = await ensureSubscription(ctx, envMap);

    expect(result.isOk()).to.be.true;
    if (result.isOk()) {
      expect(result.value.subscriptionId).to.equal("sub-123");
      expect(result.value.subscriptionName).to.equal("My Sub");
      expect(result.value.tenantId).to.equal("tenant-456");
    }
    // listSubscriptions should not have been called
    expect((ctx.auth.azureAccountProvider.listSubscriptions as sinon.SinonStub).called).to.be.false;
  });

  it("should auto-select single subscription", async () => {
    const ctx = createMockContext();
    const sub = {
      subscriptionId: "sub-only",
      subscriptionName: "Only Sub",
      tenantId: "t1",
    };
    (ctx.auth.azureAccountProvider.listSubscriptions as sinon.SinonStub).resolves([sub]);
    const envMap = new Map<string, string>();

    const result = await ensureSubscription(ctx, envMap);

    expect(result.isOk()).to.be.true;
    if (result.isOk()) {
      expect(result.value.subscriptionId).to.equal("sub-only");
    }
    // Should have set envMap
    expect(envMap.get("AZURE_SUBSCRIPTION_ID")).to.equal("sub-only");
    expect(envMap.get("AZURE_SUBSCRIPTION_NAME")).to.equal("Only Sub");
  });

  it("should prompt when multiple subscriptions exist", async () => {
    const ctx = createMockContext();
    const subs = [
      { subscriptionId: "sub-a", subscriptionName: "Sub A", tenantId: "t1" },
      { subscriptionId: "sub-b", subscriptionName: "Sub B", tenantId: "t1" },
    ];
    (ctx.auth.azureAccountProvider.listSubscriptions as sinon.SinonStub).resolves(subs);
    (ctx.ui.selectOption as sinon.SinonStub).resolves(ok({ type: "success", result: "sub-b" }));
    const envMap = new Map<string, string>();

    const result = await ensureSubscription(ctx, envMap);

    expect(result.isOk()).to.be.true;
    if (result.isOk()) {
      expect(result.value.subscriptionId).to.equal("sub-b");
      expect(result.value.subscriptionName).to.equal("Sub B");
    }
    expect(envMap.get("AZURE_SUBSCRIPTION_ID")).to.equal("sub-b");
    expect((ctx.auth.azureAccountProvider.setSubscription as sinon.SinonStub).calledWith("sub-b"))
      .to.be.true;
  });

  it("should return error when no subscriptions found", async () => {
    const ctx = createMockContext();
    (ctx.auth.azureAccountProvider.listSubscriptions as sinon.SinonStub).resolves([]);
    const envMap = new Map<string, string>();

    const result = await ensureSubscription(ctx, envMap);

    expect(result.isErr()).to.be.true;
    if (result.isErr()) {
      expect(result.error.code).to.equal("NoAzureSubscription");
    }
  });

  it("should return error when user cancels selection", async () => {
    const ctx = createMockContext();
    const subs = [
      { subscriptionId: "sub-a", subscriptionName: "Sub A", tenantId: "t1" },
      { subscriptionId: "sub-b", subscriptionName: "Sub B", tenantId: "t1" },
    ];
    (ctx.auth.azureAccountProvider.listSubscriptions as sinon.SinonStub).resolves(subs);
    (ctx.ui.selectOption as sinon.SinonStub).resolves(ok({ type: "back" }));
    const envMap = new Map<string, string>();

    const result = await ensureSubscription(ctx, envMap);

    expect(result.isErr()).to.be.true;
    if (result.isErr()) {
      expect(result.error.code).to.equal("UserCancelled");
    }
  });
});

// ---------------------------------------------------------------------------
// ensureResourceGroup
// ---------------------------------------------------------------------------

describe("ensureResourceGroup", () => {
  afterEach(() => sinon.restore());

  it("should return existing resource group from envMap", async () => {
    const ctx = createMockContext();
    const envMap = new Map([
      ["AZURE_RESOURCE_GROUP_NAME", "existing-rg"],
      ["AZURE_RESOURCE_GROUP_LOCATION", "eastus"],
    ]);

    const result = await ensureResourceGroup(ctx, envMap, "sub-1", "myapp", "dev");

    expect(result.isOk()).to.be.true;
    if (result.isOk()) {
      expect(result.value.name).to.equal("existing-rg");
      expect(result.value.location).to.equal("eastus");
      expect(result.value.isNew).to.be.false;
    }
  });

  it("should prompt for name and location when not in envMap", async () => {
    const ctx = createMockContext();
    (ctx.ui.inputText as sinon.SinonStub)
      .onFirstCall()
      .resolves(ok({ type: "success", result: "rg-test-dev" }))
      .onSecondCall()
      .resolves(ok({ type: "success", result: "westus2" }));
    const envMap = new Map<string, string>();

    const result = await ensureResourceGroup(ctx, envMap, "sub-1", "my-app", "dev");

    expect(result.isOk()).to.be.true;
    if (result.isOk()) {
      expect(result.value.name).to.equal("rg-test-dev");
      expect(result.value.location).to.equal("westus2");
      expect(result.value.isNew).to.be.true;
    }
    expect(envMap.get("AZURE_RESOURCE_GROUP_NAME")).to.equal("rg-test-dev");
    expect(envMap.get("AZURE_RESOURCE_GROUP_LOCATION")).to.equal("westus2");
  });

  it("should include RESOURCE_SUFFIX in default name", async () => {
    const ctx = createMockContext();
    (ctx.ui.inputText as sinon.SinonStub)
      .onFirstCall()
      .resolves(ok({ type: "success", result: "custom-rg" }))
      .onSecondCall()
      .resolves(ok({ type: "success", result: "centralus" }));
    const envMap = new Map<string, string>();

    await ensureResourceGroup(ctx, envMap, "sub-1", "test-app", "prod");

    // The first inputText call should have a default containing RESOURCE_SUFFIX
    const firstInputCall = (ctx.ui.inputText as sinon.SinonStub).firstCall.args[0];
    expect(firstInputCall.default).to.match(/^rg-testapp.+-prod$/);
    // RESOURCE_SUFFIX should now be in envMap (generated by ensureResourceSuffix)
    expect(envMap.get("RESOURCE_SUFFIX")).to.be.a("string").with.lengthOf(6);
  });

  it("should return error when name input is cancelled", async () => {
    const ctx = createMockContext();
    (ctx.ui.inputText as sinon.SinonStub).resolves(ok({ type: "back" }));
    const envMap = new Map<string, string>();

    const result = await ensureResourceGroup(ctx, envMap, "sub-1", "app", "dev");

    expect(result.isErr()).to.be.true;
    if (result.isErr()) {
      expect(result.error.code).to.equal("UserCancelled");
    }
  });

  it("should return error when location input is cancelled", async () => {
    const ctx = createMockContext();
    (ctx.ui.inputText as sinon.SinonStub)
      .onFirstCall()
      .resolves(ok({ type: "success", result: "rg-name" }))
      .onSecondCall()
      .resolves(ok({ type: "back" }));
    const envMap = new Map<string, string>();

    const result = await ensureResourceGroup(ctx, envMap, "sub-1", "app", "dev");

    expect(result.isErr()).to.be.true;
    if (result.isErr()) {
      expect(result.error.code).to.equal("UserCancelled");
    }
  });
});

// ---------------------------------------------------------------------------
// ensureResourceSuffix
// ---------------------------------------------------------------------------

describe("ensureResourceSuffix", () => {
  it("should generate a 6-char suffix and set it in envMap", () => {
    const envMap = new Map<string, string>();
    const suffix = ensureResourceSuffix(envMap);

    expect(suffix).to.be.a("string").with.lengthOf(6);
    expect(suffix).to.match(/^[a-z0-9]{6}$/);
    expect(envMap.get("RESOURCE_SUFFIX")).to.equal(suffix);
  });

  it("should return existing suffix without changing it", () => {
    const envMap = new Map([["RESOURCE_SUFFIX", "abc123"]]);
    const suffix = ensureResourceSuffix(envMap);

    expect(suffix).to.equal("abc123");
    expect(envMap.get("RESOURCE_SUFFIX")).to.equal("abc123");
  });
});

// ---------------------------------------------------------------------------
// confirmProvision
// ---------------------------------------------------------------------------

describe("confirmProvision", () => {
  afterEach(() => sinon.restore());

  it("should return ok when user confirms", async () => {
    const ctx = createMockContext();
    (ctx.ui.confirm as sinon.SinonStub).resolves(ok({ type: "success", result: true }));

    const result = await confirmProvision(ctx, "dev");

    expect(result.isOk()).to.be.true;
  });

  it("should include M365 and Azure info in message", async () => {
    const ctx = createMockContext();
    (ctx.ui.confirm as sinon.SinonStub).resolves(ok({ type: "success", result: true }));

    await confirmProvision(
      ctx,
      "prod",
      { tenantId: "t1", displayName: "Contoso" },
      { subscriptionId: "s1", subscriptionName: "My Sub", tenantId: "t1" }
    );

    const title = (ctx.ui.confirm as sinon.SinonStub).firstCall.args[0].title as string;
    expect(title).to.include("Contoso");
    expect(title).to.include("My Sub");
    expect(title).to.include("prod");
  });

  it("should return error when user declines", async () => {
    const ctx = createMockContext();
    (ctx.ui.confirm as sinon.SinonStub).resolves(ok({ type: "success", result: false }));

    const result = await confirmProvision(ctx, "dev");

    expect(result.isErr()).to.be.true;
    if (result.isErr()) {
      expect(result.error.code).to.equal("UserCancelled");
    }
  });

  it("should return error when confirm is not available", async () => {
    const ctx = createMockContext();
    (ctx.ui as any).confirm = undefined;

    const result = await confirmProvision(ctx, "dev");

    expect(result.isErr()).to.be.true;
    if (result.isErr()) {
      expect(result.error.code).to.equal("UserCancelled");
    }
  });
});

// ---------------------------------------------------------------------------
// confirmDeploy
// ---------------------------------------------------------------------------

describe("confirmDeploy", () => {
  afterEach(() => sinon.restore());

  it("should skip confirmation for local env", async () => {
    const ctx = createMockContext();
    const result = await confirmDeploy(ctx, "local");

    expect(result.isOk()).to.be.true;
    expect((ctx.ui.confirm as sinon.SinonStub).called).to.be.false;
  });

  it("should skip confirmation for testtool env", async () => {
    const ctx = createMockContext();
    const result = await confirmDeploy(ctx, "testtool");

    expect(result.isOk()).to.be.true;
  });

  it("should skip confirmation for playground env (case-insensitive)", async () => {
    const ctx = createMockContext();
    const result = await confirmDeploy(ctx, "Playground");

    expect(result.isOk()).to.be.true;
  });

  it("should skip confirmation for sandbox env", async () => {
    const ctx = createMockContext();
    const result = await confirmDeploy(ctx, "sandbox");

    expect(result.isOk()).to.be.true;
  });

  it("should prompt for non-local environments", async () => {
    const ctx = createMockContext();
    (ctx.ui.confirm as sinon.SinonStub).resolves(ok({ type: "success", result: true }));

    const result = await confirmDeploy(ctx, "production");

    expect(result.isOk()).to.be.true;
    expect((ctx.ui.confirm as sinon.SinonStub).calledOnce).to.be.true;
  });

  it("should return error when user cancels deploy", async () => {
    const ctx = createMockContext();
    (ctx.ui.confirm as sinon.SinonStub).resolves(ok({ type: "success", result: false }));

    const result = await confirmDeploy(ctx, "staging");

    expect(result.isErr()).to.be.true;
    if (result.isErr()) {
      expect(result.error.code).to.equal("UserCancelled");
    }
  });
});
