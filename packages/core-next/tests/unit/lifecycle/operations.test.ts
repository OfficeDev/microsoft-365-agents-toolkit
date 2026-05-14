/**
 * Copyright (c) Microsoft Corporation.
 * Licensed under the MIT license.
 */

import { expect } from "chai";
import { describe, it, beforeEach, afterEach } from "mocha";
import * as sinon from "sinon";
import { ok, err } from "neverthrow";
import * as fs from "fs";
import * as envManager from "../../../src/environment/envManager";
import * as parser from "../../../src/lifecycle/parser";
import * as executor from "../../../src/lifecycle/executor";
import * as prerequisites from "../../../src/lifecycle/prerequisites";
import * as analyzeModule from "../../../src/lifecycle/analyze";
import { provisionOp, deployOp, publishOp } from "../../../src/lifecycle/operations";
import { runOperation } from "../../../src/core/operation";
import type { AtkContext } from "../../../src/core/context";
import type { RawProjectModel, DriverStep, LifecycleResult } from "../../../src/lifecycle/types";

// ---------------------------------------------------------------------------
// Shared fixtures
// ---------------------------------------------------------------------------

const PROVISION_STEPS: DriverStep[] = [
  { uses: "teamsApp/create", with: { name: "myapp" } },
  { uses: "arm/deploy", with: { template: "./infra/main.bicep" } },
];

const DEPLOY_STEPS: DriverStep[] = [
  { uses: "azureAppService/deploy", with: { artifactFolder: "./build" } },
];

const PUBLISH_STEPS: DriverStep[] = [
  { uses: "teamsApp/publishAppPackage", with: { appPackagePath: "./build/appPackage.zip" } },
];

function makeModel(overrides?: Partial<RawProjectModel>): RawProjectModel {
  return {
    version: "v1.0",
    provision: PROVISION_STEPS,
    deploy: DEPLOY_STEPS,
    publish: PUBLISH_STEPS,
    ...overrides,
  };
}

function makeLifecycleResult(lifecycle: string, stepCount: number): LifecycleResult {
  return {
    lifecycle: lifecycle as any,
    steps: Array.from({ length: stepCount }, (_, i) => ({
      driver: `driver-${i}`,
      outputs: {},
      durationMs: 10,
    })),
    totalDurationMs: stepCount * 10,
  };
}

function createMockContext(): AtkContext {
  return {
    auth: {
      m365TokenProvider: {
        getAccessToken: sinon.stub().resolves(ok("token")),
        getJsonObject: sinon.stub().resolves(ok({ tid: "t1", name: "Test" })),
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
        listSubscriptions: sinon
          .stub()
          .resolves([{ subscriptionId: "s1", subscriptionName: "Sub", tenantId: "t1" }]),
        setSubscription: sinon.stub().resolves(),
        getAccountInfo: sinon.stub().returns({ email: "user@test.com", tenantId: "t1" }),
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
      createProgressBar: sinon
        .stub()
        .returns({ start: sinon.stub(), next: sinon.stub(), end: sinon.stub() }),
      confirm: sinon.stub(),
    },
    correlationId: "test-corr-id",
  } as any;
}

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

describe("provisionOp", () => {
  let sandbox: sinon.SinonSandbox;

  beforeEach(() => {
    sandbox = sinon.createSandbox();
  });

  afterEach(() => {
    sandbox.restore();
  });

  function stubDependencies(model: RawProjectModel) {
    sandbox.stub(fs.promises, "readFile").resolves("version: v1.0");
    sandbox.stub(parser, "parseProjectYaml").resolves(ok(model));
    sandbox.stub(envManager, "readEnvFile").resolves(ok({}));
    sandbox.stub(envManager, "writeEnvFile").resolves(ok(undefined));
  }

  it("should provision with M365 + Azure steps", async () => {
    const model = makeModel();
    stubDependencies(model);

    sandbox.stub(analyzeModule, "analyzeSteps").returns({
      needsM365: true,
      needsAzure: true,
      driverIds: ["teamsApp/create", "arm/deploy"],
      unresolvedVars: ["AZURE_RESOURCE_GROUP_NAME"],
    });

    sandbox
      .stub(prerequisites, "ensureM365Auth")
      .resolves(ok({ tenantId: "t1", displayName: "Contoso" }));
    sandbox
      .stub(prerequisites, "ensureAzureAuth")
      .resolves(ok({ accountId: "user@test.com", tenantId: "t1" }));
    sandbox.stub(prerequisites, "ensureResourceSuffix").returns("abc123");
    sandbox
      .stub(prerequisites, "ensureSubscription")
      .resolves(ok({ subscriptionId: "s1", subscriptionName: "Sub", tenantId: "t1" }));
    sandbox
      .stub(prerequisites, "ensureResourceGroup")
      .resolves(ok({ name: "rg-test", location: "centralus", isNew: true }));
    sandbox.stub(prerequisites, "confirmProvision").resolves(ok(undefined));
    sandbox.stub(executor, "executeLifecycle").resolves(ok(makeLifecycleResult("provision", 2)));

    const ctx = createMockContext();
    const result = await runOperation(provisionOp, ctx, {
      projectPath: "/project",
      envName: "dev",
    });

    expect(result.isOk()).to.be.true;
    if (result.isOk()) {
      expect(result.value.lifecycleResult.steps).to.have.lengthOf(2);
      expect(result.value.postActions.length).to.be.greaterThan(0);
      expect(result.value.postActions[0].type).to.equal("showMessage");
    }

    // Should have persisted env
    expect((envManager.writeEnvFile as sinon.SinonStub).calledOnce).to.be.true;
  });

  it("should skip consent when skipConsent is true", async () => {
    const model = makeModel();
    stubDependencies(model);

    sandbox.stub(analyzeModule, "analyzeSteps").returns({
      needsM365: false,
      needsAzure: false,
      driverIds: [],
      unresolvedVars: [],
    });

    const confirmStub = sandbox.stub(prerequisites, "confirmProvision");
    sandbox.stub(executor, "executeLifecycle").resolves(ok(makeLifecycleResult("provision", 1)));

    const ctx = createMockContext();
    await runOperation(provisionOp, ctx, {
      projectPath: "/project",
      envName: "dev",
      skipConsent: true,
    });

    expect(confirmStub.called).to.be.false;
  });

  it("should return error when YAML is missing", async () => {
    sandbox.stub(fs.promises, "readFile").rejects(new Error("ENOENT"));
    sandbox
      .stub(parser, "parseProjectYaml")
      .resolves(
        err({ code: "YamlParseError", message: "bad yaml", kind: "user", source: "test" } as any)
      );
    sandbox.stub(envManager, "readEnvFile").resolves(ok({}));

    const ctx = createMockContext();
    const result = await runOperation(provisionOp, ctx, {
      projectPath: "/missing",
      envName: "dev",
    });

    expect(result.isErr()).to.be.true;
  });

  it("should return error when no provision steps exist", async () => {
    const model = makeModel({ provision: undefined });
    stubDependencies(model);

    const ctx = createMockContext();
    const result = await runOperation(provisionOp, ctx, {
      projectPath: "/project",
      envName: "dev",
    });

    expect(result.isErr()).to.be.true;
    if (result.isErr()) {
      expect(result.error.code).to.equal("NoProvisionSteps");
    }
  });

  it("should return error when M365 auth fails", async () => {
    const model = makeModel();
    stubDependencies(model);

    sandbox.stub(analyzeModule, "analyzeSteps").returns({
      needsM365: true,
      needsAzure: false,
      driverIds: ["teamsApp/create"],
      unresolvedVars: [],
    });

    sandbox.stub(prerequisites, "ensureM365Auth").resolves(
      err({
        code: "M365AuthFailed",
        message: "Login failed",
        kind: "user",
        source: "test",
      } as any)
    );

    const ctx = createMockContext();
    const result = await runOperation(provisionOp, ctx, {
      projectPath: "/project",
      envName: "dev",
    });

    expect(result.isErr()).to.be.true;
    if (result.isErr()) {
      expect(result.error.code).to.equal("M365AuthFailed");
    }
  });

  it("should persist env even on execution failure", async () => {
    const model = makeModel();
    stubDependencies(model);

    sandbox.stub(analyzeModule, "analyzeSteps").returns({
      needsM365: false,
      needsAzure: false,
      driverIds: [],
      unresolvedVars: [],
    });

    sandbox.stub(prerequisites, "confirmProvision").resolves(ok(undefined));
    sandbox
      .stub(executor, "executeLifecycle")
      .resolves(
        err({ code: "DriverFailed", message: "step failed", kind: "system", source: "test" } as any)
      );

    const ctx = createMockContext();
    const result = await runOperation(provisionOp, ctx, {
      projectPath: "/project",
      envName: "dev",
    });

    expect(result.isErr()).to.be.true;
    // Should still have persisted partial env outputs
    expect((envManager.writeEnvFile as sinon.SinonStub).calledOnce).to.be.true;
  });

  it("should include Azure portal link post-action when Azure is used", async () => {
    const model = makeModel();
    stubDependencies(model);

    sandbox.stub(analyzeModule, "analyzeSteps").returns({
      needsM365: false,
      needsAzure: true,
      driverIds: ["arm/deploy"],
      unresolvedVars: [],
    });

    sandbox
      .stub(prerequisites, "ensureAzureAuth")
      .resolves(ok({ accountId: "user@test.com", tenantId: "t1" }));
    sandbox.stub(prerequisites, "ensureResourceSuffix").returns("abc123");
    sandbox
      .stub(prerequisites, "ensureSubscription")
      .resolves(ok({ subscriptionId: "s1", subscriptionName: "Sub", tenantId: "t1" }));
    sandbox
      .stub(prerequisites, "ensureResourceGroup")
      .resolves(ok({ name: "rg-myapp-dev", location: "centralus", isNew: false }));
    sandbox.stub(prerequisites, "confirmProvision").resolves(ok(undefined));

    // Set RG name in envMap via the executeLifecycle stub
    sandbox.stub(executor, "executeLifecycle").callsFake(async (_ctx, _lc, _steps, envMap) => {
      envMap.set("AZURE_RESOURCE_GROUP_NAME", "rg-myapp-dev");
      return ok(makeLifecycleResult("provision", 1));
    });

    const ctx = createMockContext();
    const result = await runOperation(provisionOp, ctx, {
      projectPath: "/project",
      envName: "dev",
    });

    expect(result.isOk()).to.be.true;
    if (result.isOk()) {
      const openUrlAction = result.value.postActions.find((a) => a.type === "openUrl");
      expect(openUrlAction).to.exist;
      expect(openUrlAction!.url).to.include("portal.azure.com");
      expect(openUrlAction!.url).to.include("rg-myapp-dev");
    }
  });
});

// ---------------------------------------------------------------------------
// deployOp
// ---------------------------------------------------------------------------

describe("deployOp", () => {
  let sandbox: sinon.SinonSandbox;

  beforeEach(() => {
    sandbox = sinon.createSandbox();
  });

  afterEach(() => {
    sandbox.restore();
  });

  function stubDependencies(model: RawProjectModel) {
    sandbox.stub(fs.promises, "readFile").resolves("version: v1.0");
    sandbox.stub(parser, "parseProjectYaml").resolves(ok(model));
    sandbox.stub(envManager, "readEnvFile").resolves(ok({}));
    sandbox.stub(envManager, "writeEnvFile").resolves(ok(undefined));
  }

  it("should deploy with consent confirmation", async () => {
    const model = makeModel();
    stubDependencies(model);

    sandbox.stub(prerequisites, "confirmDeploy").resolves(ok(undefined));
    sandbox.stub(executor, "executeLifecycle").resolves(ok(makeLifecycleResult("deploy", 1)));

    const ctx = createMockContext();
    const result = await runOperation(deployOp, ctx, {
      projectPath: "/project",
      envName: "prod",
    });

    expect(result.isOk()).to.be.true;
    if (result.isOk()) {
      expect(result.value.lifecycleResult.lifecycle).to.equal("deploy");
      expect(result.value.postActions[0].message).to.include("Deploy completed");
    }

    expect((prerequisites.confirmDeploy as sinon.SinonStub).calledWith(sinon.match.any, "prod")).to
      .be.true;
  });

  it("should skip consent when skipConsent is true", async () => {
    const model = makeModel();
    stubDependencies(model);

    const confirmStub = sandbox.stub(prerequisites, "confirmDeploy");
    sandbox.stub(executor, "executeLifecycle").resolves(ok(makeLifecycleResult("deploy", 1)));

    const ctx = createMockContext();
    await runOperation(deployOp, ctx, {
      projectPath: "/project",
      envName: "prod",
      skipConsent: true,
    });

    expect(confirmStub.called).to.be.false;
  });

  it("should return error when no deploy steps exist", async () => {
    const model = makeModel({ deploy: undefined });
    stubDependencies(model);

    const ctx = createMockContext();
    const result = await runOperation(deployOp, ctx, {
      projectPath: "/project",
      envName: "dev",
    });

    expect(result.isErr()).to.be.true;
    if (result.isErr()) {
      expect(result.error.code).to.equal("NoDeploySteps");
    }
  });

  it("should return error when user cancels deploy consent", async () => {
    const model = makeModel();
    stubDependencies(model);

    sandbox
      .stub(prerequisites, "confirmDeploy")
      .resolves(
        err({ code: "UserCancelled", message: "cancelled", kind: "user", source: "test" } as any)
      );

    const ctx = createMockContext();
    const result = await runOperation(deployOp, ctx, {
      projectPath: "/project",
      envName: "staging",
    });

    expect(result.isErr()).to.be.true;
    if (result.isErr()) {
      expect(result.error.code).to.equal("UserCancelled");
    }
  });

  it("should persist env on failure", async () => {
    const model = makeModel();
    stubDependencies(model);

    sandbox.stub(prerequisites, "confirmDeploy").resolves(ok(undefined));
    sandbox
      .stub(executor, "executeLifecycle")
      .resolves(
        err({ code: "DeployFailed", message: "error", kind: "system", source: "test" } as any)
      );

    const ctx = createMockContext();
    await runOperation(deployOp, ctx, {
      projectPath: "/project",
      envName: "dev",
    });

    expect((envManager.writeEnvFile as sinon.SinonStub).calledOnce).to.be.true;
  });
});

// ---------------------------------------------------------------------------
// publishOp
// ---------------------------------------------------------------------------

describe("publishOp", () => {
  let sandbox: sinon.SinonSandbox;

  beforeEach(() => {
    sandbox = sinon.createSandbox();
  });

  afterEach(() => {
    sandbox.restore();
  });

  function stubDependencies(model: RawProjectModel) {
    sandbox.stub(fs.promises, "readFile").resolves("version: v1.0");
    sandbox.stub(parser, "parseProjectYaml").resolves(ok(model));
    sandbox.stub(envManager, "readEnvFile").resolves(ok({}));
    sandbox.stub(envManager, "writeEnvFile").resolves(ok(undefined));
  }

  it("should publish and include Teams Admin Center link", async () => {
    const model = makeModel();
    stubDependencies(model);

    sandbox.stub(executor, "executeLifecycle").resolves(ok(makeLifecycleResult("publish", 1)));

    const ctx = createMockContext();
    const result = await runOperation(publishOp, ctx, {
      projectPath: "/project",
      envName: "prod",
    });

    expect(result.isOk()).to.be.true;
    if (result.isOk()) {
      expect(result.value.lifecycleResult.lifecycle).to.equal("publish");
      const adminLink = result.value.postActions.find(
        (a) => a.type === "openUrl" && a.url?.includes("admin.teams.microsoft.com")
      );
      expect(adminLink).to.exist;
    }
  });

  it("should return error when no publish steps exist", async () => {
    const model = makeModel({ publish: undefined });
    stubDependencies(model);

    const ctx = createMockContext();
    const result = await runOperation(publishOp, ctx, {
      projectPath: "/project",
      envName: "dev",
    });

    expect(result.isErr()).to.be.true;
    if (result.isErr()) {
      expect(result.error.code).to.equal("NoPublishSteps");
    }
  });

  it("should persist env on failure", async () => {
    const model = makeModel();
    stubDependencies(model);

    sandbox
      .stub(executor, "executeLifecycle")
      .resolves(
        err({ code: "PublishFailed", message: "error", kind: "system", source: "test" } as any)
      );

    const ctx = createMockContext();
    await runOperation(publishOp, ctx, {
      projectPath: "/project",
      envName: "dev",
    });

    expect((envManager.writeEnvFile as sinon.SinonStub).calledOnce).to.be.true;
  });

  it("should reject invalid input", async () => {
    const ctx = createMockContext();
    const result = await runOperation(publishOp, ctx, {
      projectPath: "",
      envName: "dev",
    });

    expect(result.isErr()).to.be.true;
    if (result.isErr()) {
      expect(result.error.code).to.equal("InputValidationError");
    }
  });
});
