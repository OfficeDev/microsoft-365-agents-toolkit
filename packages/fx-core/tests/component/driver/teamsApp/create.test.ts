// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { err, ok, UserError } from "@microsoft/teamsfx-api";
import AdmZip from "adm-zip";
import fs from "fs-extra";
import mockedEnv from "mocked-env";
import { v4 as uuid } from "uuid";
import isUUID from "validator/lib/isUUID";
import { teamsDevPortalClient } from "../../../../src/client/teamsDevPortalClientProvider";
import { SovereignCloudEnvironment } from "../../../../src/common/accountUtils";
import { FeatureFlagName } from "../../../../src/common/featureFlags";
import { CreateTeamsAppDriver } from "../../../../src/component/driver/teamsApp/create";
import { CreateTeamsAppArgs } from "../../../../src/component/driver/teamsApp/interfaces/CreateTeamsAppArgs";
import { MockedLogProvider, MockedUserInteraction } from "../../../plugins/solution/util";
import { AppDefinition } from "./../../../../src/component/driver/teamsApp/interfaces/appdefinitions/appDefinition";
import { MockedM365Provider } from "../../../core/utils";
import { chai, expect, vi } from "vitest";

describe("teamsApp/create", async () => {
  const teamsAppDriver = new CreateTeamsAppDriver();
  let restoreEnv: (() => void) | undefined;
  const mockedDriverContext: any = {
    m365TokenProvider: new MockedM365Provider(),
    logProvider: new MockedLogProvider(),
    ui: new MockedUserInteraction(),
    projectPath: "./",
  };

  const appId = uuid();
  const appDef: AppDefinition = {
    appId,
    appName: "fake",
    teamsAppId: appId,
    userList: [],
    tenantId: uuid(),
  };
  const nativeApp = {
    appId,
    appProfile: { appDetails: { applicationManifest: { id: appId } } },
  };

  beforeEach(() => {
    process.env[FeatureFlagName.NewDeveloperPortalApis] = "true";
    vi.spyOn(mockedDriverContext.m365TokenProvider, "getJsonObject").mockResolvedValue(
      ok({ tid: appDef.tenantId })
    );
  });

  afterEach(() => {
    vi.restoreAllMocks();
    delete process.env[FeatureFlagName.NewDeveloperPortalApis];
    restoreEnv?.();
    restoreEnv = undefined;
  });

  it("skip create in GCCH", async () => {
    restoreEnv = mockedEnv({
      [FeatureFlagName.SovereignCloudEnvironment]: SovereignCloudEnvironment.GCCH,
    });
    const createAppSpy = vi.spyOn(teamsDevPortalClient, "createApp");
    const readFileStub = vi.spyOn(fs, "readFile");

    const args: CreateTeamsAppArgs = {
      name: appDef.appName!,
    };

    const result = (await teamsAppDriver.execute(args, mockedDriverContext)).result;
    chai.assert(result.isOk());
    expect(createAppSpy).not.toHaveBeenCalled();
    expect(readFileStub).not.toHaveBeenCalled();
  });

  it("skip create in DoD", async () => {
    restoreEnv = mockedEnv({
      [FeatureFlagName.SovereignCloudEnvironment]: SovereignCloudEnvironment.DOD,
    });
    const createAppSpy = vi.spyOn(teamsDevPortalClient, "createApp");
    const readFileStub = vi.spyOn(fs, "readFile");

    const args: CreateTeamsAppArgs = {
      name: appDef.appName!,
    };

    const result = (await teamsAppDriver.execute(args, mockedDriverContext)).result;
    chai.assert(result.isOk());
    expect(createAppSpy).not.toHaveBeenCalled();
    expect(readFileStub).not.toHaveBeenCalled();
  });

  it("invalid param error", async () => {
    const args: CreateTeamsAppArgs = {
      name: "",
    };
    const result = (await teamsAppDriver.execute(args, mockedDriverContext)).result;
    chai.assert(result.isErr());
    if (result.isErr()) {
      chai.assert.equal("InvalidActionInputError", result.error.name);
    }
  });

  it("happy path", async () => {
    process.env[FeatureFlagName.NewDeveloperPortalApis] = "true";
    const args: CreateTeamsAppArgs = {
      name: appDef.appName!,
    };

    vi.spyOn(teamsDevPortalClient, "getApp").mockImplementation(() => {
      throw new Error("404");
    });
    const createAppSpy = vi.spyOn(teamsDevPortalClient, "createApp").mockResolvedValue(nativeApp);
    const readFileSpy = vi.spyOn(fs, "readFile");

    const result = (await teamsAppDriver.execute(args, mockedDriverContext)).result;
    chai.assert.isTrue(result.isOk());
    expect(createAppSpy).toHaveBeenCalledWith("fakeToken", appDef.appName);
    expect(readFileSpy).not.toHaveBeenCalled();
    if (result.isOk()) {
      chai.assert.equal(result.value.get("APP_ID"), appId);
    }
  });

  it("uses legacy app package import when the new API is disabled", async () => {
    process.env[FeatureFlagName.NewDeveloperPortalApis] = "false";
    const args: CreateTeamsAppArgs = { name: appDef.appName! };
    const importAppSpy = vi.spyOn(teamsDevPortalClient, "importApp").mockResolvedValue(appDef);
    const createAppSpy = vi.spyOn(teamsDevPortalClient, "createApp");

    const result = (await teamsAppDriver.execute(args, mockedDriverContext)).result;

    chai.assert.isTrue(result.isOk());
    expect(importAppSpy).toHaveBeenCalledOnce();
    expect(createAppSpy).not.toHaveBeenCalled();
  });

  it.each(["appId", "teamsAppId"])(
    "AC-20: flag-off create honors explicit %s mapping without canonical fallback",
    async (key) => {
      restoreEnv = mockedEnv({
        [FeatureFlagName.NewDeveloperPortalApis]: "false",
        CUSTOM_ID: "legacy-id",
        APP_ID: uuid(),
        TEAMS_APP_ID: uuid(),
        APP_MANIFEST_ID: undefined,
      });
      const get = vi
        .spyOn(teamsDevPortalClient, "getApp")
        .mockRejectedValue(new Error("not found"));
      const imported = vi.spyOn(teamsDevPortalClient, "importApp").mockResolvedValue(appDef);
      const result = (
        await teamsAppDriver.execute(
          { name: "legacy" },
          mockedDriverContext,
          new Map([[key, "CUSTOM_ID"]])
        )
      ).result._unsafeUnwrap();
      expect(get).toHaveBeenCalledWith("fakeToken", "legacy-id");
      expect(imported).toHaveBeenCalledOnce();
      expect(result.get("CUSTOM_ID")).toBe(appDef.teamsAppId);
      expect(result.has("APP_ID")).toBe(false);
    }
  );

  it("generates a manifest app ID when the environment value is empty", async () => {
    process.env[FeatureFlagName.NewDeveloperPortalApis] = "false";
    restoreEnv = mockedEnv({ TEAMS_APP_ID: "" });
    const importAppSpy = vi.spyOn(teamsDevPortalClient, "importApp").mockResolvedValue(appDef);

    const result = (await teamsAppDriver.execute({ name: appDef.appName! }, mockedDriverContext))
      .result;

    chai.assert.isTrue(result.isOk());
    const zip = new AdmZip(importAppSpy.mock.calls[0][1]);
    const manifest = JSON.parse(zip.readAsText("manifest.json"));
    chai.assert.isTrue(isUUID(manifest.id));
  });

  it("app exists", async () => {
    const args: CreateTeamsAppArgs = {
      name: appDef.appName!,
    };

    restoreEnv = mockedEnv({ TEAMS_APP_ID: appId });
    vi.spyOn(teamsDevPortalClient, "getApp").mockResolvedValue(nativeApp);
    const createAppSpy = vi.spyOn(teamsDevPortalClient, "createApp");

    const result = (await teamsAppDriver.execute(args, mockedDriverContext)).result;
    chai.assert.isTrue(result.isOk());
    expect(createAppSpy).not.toHaveBeenCalled();
  });

  it("AC-14: preserves both identities when the Developer Portal resource ID differs", async () => {
    const manifestId = uuid();
    const resourceAppId = uuid();
    restoreEnv = mockedEnv({ APP_ID: resourceAppId, APP_MANIFEST_ID: manifestId });
    vi.spyOn(teamsDevPortalClient, "getApp").mockResolvedValue({
      appId: resourceAppId,
      appProfile: { appDetails: { applicationManifest: { id: manifestId } } },
    });

    const result = (await teamsAppDriver.execute({ name: appDef.appName! }, mockedDriverContext))
      .result;

    chai.assert.isTrue(result.isOk());
    if (result.isOk()) {
      chai.assert.equal(result.value.get("APP_ID"), resourceAppId);
      chai.assert.equal(process.env.APP_MANIFEST_ID, manifestId);
    }
  });

  it.each([
    { responseId: "", savedId: undefined, manifestId: appId, savedManifestId: undefined },
    { responseId: "invalid", savedId: undefined, manifestId: appId, savedManifestId: undefined },
    { responseId: uuid(), savedId: appId, manifestId: appId, savedManifestId: undefined },
    { responseId: appId, savedId: appId, manifestId: uuid(), savedManifestId: undefined },
    { responseId: appId, savedId: appId, manifestId: appId, savedManifestId: uuid() },
  ])("AC-04/AC-18: rejects inconsistent resource or manifest identity (%j)", async (test) => {
    restoreEnv = mockedEnv({
      APP_ID: test.savedId,
      TEAMS_APP_ID: undefined,
      APP_MANIFEST_ID: test.savedManifestId,
    });
    const response = {
      appId: test.responseId,
      appProfile: { appDetails: { applicationManifest: { id: test.manifestId } } },
    };
    const create = vi.spyOn(teamsDevPortalClient, "createApp").mockResolvedValue(response);
    vi.spyOn(teamsDevPortalClient, "getApp").mockResolvedValue(response);
    const result = (await teamsAppDriver.execute({ name: "identity" }, mockedDriverContext)).result;
    expect(result.isErr()).toBe(true);
    if (test.savedId) expect(create).not.toHaveBeenCalled();
  });

  it("AC-04: invalid selected state fails before any DP mutation", async () => {
    restoreEnv = mockedEnv({ APP_ID: "invalid", TEAMS_APP_ID: appId });
    const create = vi.spyOn(teamsDevPortalClient, "createApp");
    const get = vi.spyOn(teamsDevPortalClient, "getApp");
    const result = (await teamsAppDriver.execute({ name: "identity" }, mockedDriverContext)).result;
    expect(result.isErr()).toBe(true);
    expect(create).not.toHaveBeenCalled();
    expect(get).not.toHaveBeenCalled();
  });

  it("AC-10: propagates authenticated tenant lookup failure without writing outputs", async () => {
    const failure = new UserError("test", "TenantLookupFailed", "tenant unavailable");
    vi.mocked(mockedDriverContext.m365TokenProvider.getJsonObject).mockResolvedValue(err(failure));
    vi.spyOn(teamsDevPortalClient, "createApp").mockResolvedValue(nativeApp);
    const result = (await teamsAppDriver.execute({ name: "identity" }, mockedDriverContext)).result;
    expect(result.isErr()).toBe(true);
    if (result.isErr()) expect(result.error).toBe(failure);
  });

  it("AC-10: legacy responses without a tenant cannot produce successful outputs", async () => {
    process.env[FeatureFlagName.NewDeveloperPortalApis] = "false";
    restoreEnv = mockedEnv({ TEAMS_APP_ID: appId });
    vi.spyOn(teamsDevPortalClient, "getApp").mockResolvedValue({ teamsAppId: appId });
    const result = (await teamsAppDriver.execute({ name: "identity" }, mockedDriverContext)).result;
    expect(result.isErr()).toBe(true);
  });

  it("does not create a replacement when existing app lookup fails", async () => {
    const args: CreateTeamsAppArgs = {
      name: appDef.appName!,
    };

    restoreEnv = mockedEnv({ TEAMS_APP_ID: appId });
    vi.spyOn(teamsDevPortalClient, "getApp").mockRejectedValue(new Error("lookup failed"));
    const createAppSpy = vi.spyOn(teamsDevPortalClient, "createApp");

    const result = (await teamsAppDriver.execute(args, mockedDriverContext)).result;

    chai.assert.isTrue(result.isErr());
    expect(createAppSpy).not.toHaveBeenCalled();
  });

  it("preserves a typed error when existing app lookup fails", async () => {
    restoreEnv = mockedEnv({ TEAMS_APP_ID: appId });
    const lookupError = new UserError("source", "LookupFailed", "lookup failed", "lookup failed");
    vi.spyOn(teamsDevPortalClient, "getApp").mockRejectedValue(lookupError);

    const result = (await teamsAppDriver.execute({ name: appDef.appName! }, mockedDriverContext))
      .result;

    chai.assert.isTrue(result.isErr());
    if (result.isErr()) {
      chai.assert.strictEqual(result.error, lookupError);
    }
  });

  it("reuses the created app when provision is retried", async () => {
    const args: CreateTeamsAppArgs = { name: appDef.appName! };
    const createAppSpy = vi.spyOn(teamsDevPortalClient, "createApp").mockResolvedValue(nativeApp);
    const getAppSpy = vi.spyOn(teamsDevPortalClient, "getApp").mockResolvedValue(nativeApp);

    const firstResult = (await teamsAppDriver.execute(args, mockedDriverContext)).result;
    chai.assert.isTrue(firstResult.isOk());

    restoreEnv = mockedEnv({ TEAMS_APP_ID: appId });
    const retryResult = (await teamsAppDriver.execute(args, mockedDriverContext)).result;

    chai.assert.isTrue(retryResult.isOk());
    expect(createAppSpy).toHaveBeenCalledOnce();
    expect(getAppSpy).toHaveBeenCalledWith("fakeToken", appId);
  });

  it("API failure", async () => {
    const args: CreateTeamsAppArgs = {
      name: appDef.appName!,
    };
    vi.spyOn(teamsDevPortalClient, "getApp").mockImplementation(() => {
      throw new Error("404");
    });
    vi.spyOn(teamsDevPortalClient, "createApp").mockImplementation(() => {
      throw new Error("409");
    });

    const result = (await teamsAppDriver.execute(args, mockedDriverContext)).result;
    chai.assert.isTrue(result.isErr());
  });

  it("Token error", async () => {
    const args: CreateTeamsAppArgs = {
      name: appDef.appName!,
    };
    vi.spyOn(MockedM365Provider.prototype, "getAccessToken").mockResolvedValue(
      err(new UserError({}))
    );
    const result = (await teamsAppDriver.execute(args, mockedDriverContext)).result;
    chai.assert.isTrue(result.isErr());
  });

  it("respects user-configured teamsAppTenantId env var name", async () => {
    const args: CreateTeamsAppArgs = {
      name: appDef.appName!,
    };

    vi.spyOn(teamsDevPortalClient, "getApp").mockImplementation(() => {
      throw new Error("404");
    });
    vi.spyOn(teamsDevPortalClient, "createApp").mockResolvedValue(nativeApp);

    const outputEnvVarNames = new Map<string, string>([
      ["teamsAppId", "MY_TEAMS_APP_ID"],
      ["teamsAppTenantId", "MY_TEAMS_APP_TENANT_ID"],
    ]);

    const result = (await teamsAppDriver.execute(args, mockedDriverContext, outputEnvVarNames))
      .result;
    chai.assert.isTrue(result.isOk());
    if (result.isOk()) {
      chai.assert.equal(result.value.get("MY_TEAMS_APP_ID"), appDef.appId);
      chai.assert.equal(result.value.get("MY_TEAMS_APP_TENANT_ID"), appDef.tenantId);
      // The internal default name must not leak through when the author
      // configured a custom env var name.
      chai.assert.isFalse(result.value.has("TEAMS_APP_TENANT_ID"));
    }
  });

  it("falls back to TEAMS_APP_TENANT_ID when teamsAppTenantId is not configured", async () => {
    const args: CreateTeamsAppArgs = {
      name: appDef.appName!,
    };

    vi.spyOn(teamsDevPortalClient, "getApp").mockImplementation(() => {
      throw new Error("404");
    });
    vi.spyOn(teamsDevPortalClient, "createApp").mockResolvedValue(nativeApp);

    const outputEnvVarNames = new Map<string, string>([["teamsAppId", "TEAMS_APP_ID"]]);

    const result = (await teamsAppDriver.execute(args, mockedDriverContext, outputEnvVarNames))
      .result;
    chai.assert.isTrue(result.isOk());
    if (result.isOk()) {
      chai.assert.equal(result.value.get("TEAMS_APP_ID"), appDef.appId);
      chai.assert.equal(result.value.get("TEAMS_APP_TENANT_ID"), appDef.tenantId);
    }
  });

  it("AC-02: writes the canonical output to a custom destination", async () => {
    vi.spyOn(teamsDevPortalClient, "createApp").mockResolvedValue(nativeApp);
    const result = (
      await teamsAppDriver.execute(
        { name: "example" },
        mockedDriverContext,
        new Map([["appId", "MY_APP"]])
      )
    ).result;
    expect(result.isOk()).toBe(true);
    if (result.isOk()) {
      expect(result.value.get("MY_APP")).toBe(appId);
      expect(result.value.has("TEAMS_APP_ID")).toBe(false);
      expect(result.value.has("APP_MANIFEST_ID")).toBe(false);
    }
  });

  it("AC-04: rejects both output names before calling DP", async () => {
    const create = vi.spyOn(teamsDevPortalClient, "createApp");
    const result = (
      await teamsAppDriver.execute(
        { name: "example" },
        mockedDriverContext,
        new Map([
          ["appId", "APP_ID"],
          ["teamsAppId", "TEAMS_APP_ID"],
        ])
      )
    ).result;
    expect(result.isErr()).toBe(true);
    expect(create).not.toHaveBeenCalled();
  });
});
