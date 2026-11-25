// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { Platform, err, ok } from "@microsoft/teamsfx-api";
import axios, { AxiosAdapter } from "axios";
import { existsSync, mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { afterEach, describe, expect, it, vi } from "vitest";
import { teamsappMgr } from "../../src/component/driver/teamsApp/teamsappMgr";
import { PackageService } from "../../src/component/m365/packageService";
import { envUtil } from "../../src/component/utils/envUtil";
import { FxCore } from "../../src/core/FxCore";
import { FxCoreClient } from "../../src/core/FxCoreClient";
import { UserCancelError } from "../../src/error";
import {
  createAppPackage,
  createSppkg,
  productId,
  tenantId,
  tenantUrl,
} from "../component/driver/deploy/spfx/preparedPublicationFixtures";
import { MockTools } from "./utils";

describe("FxCoreClient", () => {
  const originalAdapter = axios.defaults.adapter;
  const originalRaiValidationEnabled = process.env.TEAMSFX_RAI_VALIDATION_ENABLED;

  afterEach(() => {
    axios.defaults.adapter = originalAdapter;
    if (originalRaiValidationEnabled === undefined) {
      delete process.env.TEAMSFX_RAI_VALIDATION_ENABLED;
    } else {
      process.env.TEAMSFX_RAI_VALIDATION_ENABLED = originalRaiValidationEnabled;
    }
    vi.restoreAllMocks();
  });

  it("addOperations_PreAbortedSignal_DoNotInvokeEngine", async () => {
    const addPluginSpy = vi.spyOn(FxCore.prototype, "addPlugin");
    const addSkillSpy = vi.spyOn(FxCore.prototype, "addSkill");
    const addAuthActionSpy = vi.spyOn(FxCore.prototype, "addAuthAction");
    const controller = new AbortController();
    controller.abort();
    const client = new FxCoreClient(new MockTools());
    const addPluginInputs = {
      platform: Platform.CLI,
      projectPath: "project",
      "manifest-path": "project/appPackage/manifest.json",
      "api-plugin-type": "api-spec" as const,
      "openapi-spec-type": "open-file" as const,
      "openapi-spec-location": "openapi.yaml",
    };
    const addSkillInputs = {
      platform: Platform.CLI,
      projectPath: "project",
      "manifest-path": "project/appPackage/manifest.json",
      "skill-name": "weather",
      "skill-description": "Weather instructions",
    };
    const addAuthInputs = {
      platform: Platform.CLI,
      projectPath: "project",
      "plugin-manifest-path": "project/appPackage/ai-plugin.json",
      "auth-name": "apiAuth",
      "api-auth": "bearer-token" as const,
    };

    const results = await Promise.all([
      client.addPlugin(addPluginInputs, { signal: controller.signal }),
      client.addSkill(addSkillInputs, { signal: controller.signal }),
      client.addAuthAction(addAuthInputs, { signal: controller.signal }),
    ]);

    results.forEach((result) => {
      expect(result.isErr()).toBe(true);
      if (result.isErr()) expect(result.error.name).toBe("UserCancel");
    });
    expect(addPluginSpy).not.toHaveBeenCalled();
    expect(addSkillSpy).not.toHaveBeenCalled();
    expect(addAuthActionSpy).not.toHaveBeenCalled();
  });

  it("addOperations_ActiveSignal_IsForwardedToEveryEngineInput", async () => {
    const addPluginSpy = vi.spyOn(FxCore.prototype, "addPlugin").mockResolvedValue(ok(undefined));
    const addSkillSpy = vi.spyOn(FxCore.prototype, "addSkill").mockResolvedValue(ok(undefined));
    const addAuthActionSpy = vi
      .spyOn(FxCore.prototype, "addAuthAction")
      .mockResolvedValue(ok(undefined));
    const controller = new AbortController();
    const client = new FxCoreClient(new MockTools());
    const addPluginInputs = {
      platform: Platform.CLI,
      projectPath: "project",
      "manifest-path": "project/appPackage/manifest.json",
      "api-plugin-type": "api-spec" as const,
      "openapi-spec-type": "open-file" as const,
      "openapi-spec-location": "openapi.yaml",
    };
    const addSkillInputs = {
      platform: Platform.CLI,
      projectPath: "project",
      "manifest-path": "project/appPackage/manifest.json",
      "skill-from": "skills/weather",
    };
    const addAuthInputs = {
      platform: Platform.CLI,
      projectPath: "project",
      "plugin-manifest-path": "project/appPackage/ai-plugin.json",
      "auth-name": "apiAuth",
      "api-auth": "api-key" as const,
      "api-key-in": "header" as const,
      "api-key-name": "x-api-key",
    };

    const results = await Promise.all([
      client.addPlugin(addPluginInputs, { signal: controller.signal }),
      client.addSkill(addSkillInputs, { signal: controller.signal }),
      client.addAuthAction(addAuthInputs, { signal: controller.signal }),
    ]);

    results.forEach((result) => expect(result.isOk()).toBe(true));
    [addPluginSpy, addSkillSpy, addAuthActionSpy].forEach((spy) => {
      expect(spy).toHaveBeenCalledWith(expect.objectContaining({ abortSignal: controller.signal }));
    });
  });

  it("addPlugin_AbortedAfterCommitStarts_ReturnsEngineResult", async () => {
    const controller = new AbortController();
    vi.spyOn(FxCore.prototype, "addPlugin").mockImplementation(async () => {
      controller.abort();
      return ok(undefined);
    });
    const client = new FxCoreClient(new MockTools());

    const result = await client.addPlugin(
      {
        platform: Platform.CLI,
        projectPath: "project",
        "manifest-path": "project/appPackage/manifest.json",
        "api-plugin-type": "api-spec",
        "openapi-spec-type": "open-file",
        "openapi-spec-location": "openapi.yaml",
      },
      { signal: controller.signal }
    );

    expect(result.isOk()).toBe(true);
  });

  it("addAuthAction_EngineFailure_ReturnsOriginalError", async () => {
    const expectedError = new UserCancelError("test");
    vi.spyOn(FxCore.prototype, "addAuthAction").mockResolvedValue(err(expectedError));
    const client = new FxCoreClient(new MockTools());

    const result = await client.addAuthAction({
      platform: Platform.CLI,
      projectPath: "project",
      "plugin-manifest-path": "project/appPackage/ai-plugin.json",
      "auth-name": "apiAuth",
    });

    expect(result.isErr()).toBe(true);
    if (result.isErr()) expect(result.error).toBe(expectedError);
  });

  it("provision_PreAbortedSignal_DoesNotInvokeEngine", async () => {
    const provisionSpy = vi.spyOn(FxCore.prototype, "provisionResources");
    const controller = new AbortController();
    controller.abort();
    const client = new FxCoreClient(new MockTools());

    const result = await client.provision(
      { platform: Platform.CLI, projectPath: "project", env: "dev" },
      { signal: controller.signal }
    );

    expect(result.isErr()).toBe(true);
    if (result.isErr()) {
      expect(result.error.name).toBe("UserCancel");
    }
    expect(provisionSpy).not.toHaveBeenCalled();
  });

  it("provision_CompletedOperation_ReturnsEnvironmentOutputs", async () => {
    const controller = new AbortController();
    const provisionSpy = vi.spyOn(FxCore.prototype, "provisionResources").mockResolvedValue(ok());
    vi.spyOn(envUtil, "readEnv").mockResolvedValue(ok({ APP_ID: "app-id" }));
    const client = new FxCoreClient(new MockTools());

    const result = await client.provision(
      { platform: Platform.CLI, projectPath: "project", env: "dev" },
      { signal: controller.signal }
    );

    expect(result.isOk()).toBe(true);
    if (result.isOk()) {
      expect(result.value.outputs).toEqual({ APP_ID: "app-id" });
    }
    expect(provisionSpy).toHaveBeenCalledWith(
      expect.objectContaining({ abortSignal: controller.signal })
    );
  });

  it("package_CompletedOperation_ReturnsProducedPath", async () => {
    vi.spyOn(FxCore.prototype, "packageTeamsAppCLIV3").mockImplementation(async (inputs) => {
      inputs["output-package-file"] = "project/appPackage/build/appPackage.dev.zip";
      return ok(undefined);
    });
    const client = new FxCoreClient(new MockTools());

    const result = await client.package({
      platform: Platform.CLI,
      projectPath: "project",
      env: "dev",
    });

    expect(result.isOk()).toBe(true);
    if (result.isOk()) {
      expect(result.value.packagePath).toBe("project/appPackage/build/appPackage.dev.zip");
    }
  });

  it("package_AbortedDuringOperation_ReturnsUserCancel", async () => {
    const controller = new AbortController();
    vi.spyOn(FxCore.prototype, "packageTeamsAppCLIV3").mockImplementation(async () => {
      controller.abort();
      return ok();
    });
    const client = new FxCoreClient(new MockTools());

    const result = await client.package(
      { platform: Platform.CLI, projectPath: "project", env: "dev" },
      { signal: controller.signal }
    );

    expect(result.isErr()).toBe(true);
    if (result.isErr()) {
      expect(result.error.name).toBe("UserCancel");
    }
  });

  it("PUBLISH-CLIENT-02: publish forwards inputs and signal to the YAML lifecycle", async () => {
    const publishSpy = vi.spyOn(FxCore.prototype, "publishApplication").mockResolvedValue(ok());
    const legacyPublishSpy = vi.spyOn(FxCore.prototype, "publishTeamsAppCLIV3");
    const outputs = {
      M365_PUBLISHED_TITLE_ID: "published-title-id",
      M365_PUBLISHED_APP_ID: "published-app-id",
      CUSTOM_PUBLISH_OUTPUT: "custom-value",
    };
    const readEnvSpy = vi.spyOn(envUtil, "readEnv").mockResolvedValue(ok(outputs));
    const controller = new AbortController();
    const client = new FxCoreClient(new MockTools());
    const inputs = {
      platform: Platform.CLI,
      nonInteractive: true,
      projectPath: "project",
      env: "dev",
    };
    const result = await client.publish(inputs, { signal: controller.signal });

    expect(result.isOk()).toBe(true);
    if (result.isOk()) {
      expect(result.value).toEqual({ outputs });
    }
    expect(readEnvSpy).toHaveBeenCalledWith("project", "dev", false, false);
    expect(publishSpy).toHaveBeenCalledWith({ ...inputs, abortSignal: controller.signal });
    expect(legacyPublishSpy).not.toHaveBeenCalled();
    expect(inputs).not.toHaveProperty("abortSignal");
  });

  it("PUBLISH-CLIENT-03: pre-cancelled publish never invokes the engine", async () => {
    const publishSpy = vi.spyOn(FxCore.prototype, "publishApplication");
    const legacyPublishSpy = vi.spyOn(FxCore.prototype, "publishTeamsAppCLIV3");
    const controller = new AbortController();
    controller.abort();
    const client = new FxCoreClient(new MockTools());

    const result = await client.publish(
      { platform: Platform.CLI, projectPath: "project", env: "dev" },
      { signal: controller.signal }
    );

    expect(result.isErr()).toBe(true);
    if (result.isErr()) expect(result.error.name).toBe("UserCancel");
    expect(publishSpy).not.toHaveBeenCalled();
    expect(legacyPublishSpy).not.toHaveBeenCalled();
  });

  it("PUBLISH-CLIENT-03: cancellation during successful publishing returns UserCancel", async () => {
    const controller = new AbortController();
    vi.spyOn(FxCore.prototype, "publishApplication").mockImplementation(async () => {
      controller.abort();
      return ok(undefined);
    });
    const client = new FxCoreClient(new MockTools());

    const result = await client.publish(
      { platform: Platform.CLI, projectPath: "project", env: "dev" },
      { signal: controller.signal }
    );

    expect(result.isErr()).toBe(true);
    if (result.isErr()) expect(result.error.name).toBe("UserCancel");
  });

  it.each([false, true])(
    "PUBLISH-CLIENT-04: lifecycle errors are preserved with cancellation=%s",
    async (cancel) => {
      const expectedError = new UserCancelError("test-publish");
      const controller = new AbortController();
      vi.spyOn(FxCore.prototype, "publishApplication").mockImplementation(async () => {
        if (cancel) controller.abort();
        return err(expectedError);
      });
      const client = new FxCoreClient(new MockTools());

      const result = await client.publish(
        { platform: Platform.CLI, projectPath: "project", env: "dev" },
        { signal: controller.signal }
      );

      expect(result.isErr()).toBe(true);
      if (result.isErr()) expect(result.error).toBe(expectedError);
    }
  );

  it("PUBLISH-CLIENT-05: environment read errors are preserved after publishing", async () => {
    vi.spyOn(FxCore.prototype, "publishApplication").mockResolvedValue(ok());
    const expectedError = new UserCancelError("test-read-env");
    vi.spyOn(envUtil, "readEnv").mockResolvedValue(err(expectedError));
    const client = new FxCoreClient(new MockTools());

    const result = await client.publish({
      platform: Platform.CLI,
      projectPath: "project",
      env: "dev",
    });

    expect(result.isErr()).toBe(true);
    if (result.isErr()) expect(result.error).toBe(expectedError);
  });

  it.each(["dev", "local"])(
    "PUBLISH-CLIENT-01: publishing %s executes real YAML without tenant writes",
    async (env) => {
      const folder = mkdtempSync(join(tmpdir(), "fx-core-client-publish-"));
      try {
        mkdirSync(join(folder, "env"));
        writeFileSync(join(folder, "env", `.env.${env}`), `TEAMSFX_ENV=${env}\n`);
        writeFileSync(join(folder, "m365agents.yml"), "version: v1.12\n");
        writeFileSync(
          join(folder, env === "local" ? "m365agents.local.yml" : "m365agents.yml"),
          [
            "version: v1.12",
            "environmentFolderPath: ./env",
            "provision:",
            "  - uses: file/createOrUpdateEnvironmentFile",
            "    with:",
            "      target: ./provision.env",
            "      envs:",
            '        PROVISION_EXECUTED: "true"',
            "publish:",
            "  - uses: file/createOrUpdateEnvironmentFile",
            "    with:",
            `      target: ./env/.env.${env}`,
            "      envs:",
            `        PUBLISHED_ENV: "${env}"`,
            '        M365_PUBLISHED_TITLE_ID: "published-title-id"',
            '        M365_PUBLISHED_APP_ID: "published-app-id"',
            "",
          ].join("\n")
        );
        const client = new FxCoreClient(new MockTools());

        const result = await client.publish({
          platform: Platform.CLI,
          nonInteractive: true,
          projectPath: folder,
          env,
        });

        expect(result.isOk()).toBe(true);
        if (result.isOk()) {
          expect(result.value).toEqual({
            outputs: {
              TEAMSFX_ENV: env,
              PUBLISHED_ENV: env,
              M365_PUBLISHED_TITLE_ID: "published-title-id",
              M365_PUBLISHED_APP_ID: "published-app-id",
            },
          });
        }
        expect(readFileSync(join(folder, "env", `.env.${env}`), "utf8")).toContain(
          `PUBLISHED_ENV=${env}`
        );
        expect(existsSync(join(folder, "provision.env"))).toBe(false);
      } finally {
        rmSync(folder, { recursive: true, force: true });
      }
    }
  );

  it.each([false, true])(
    "SPFX-PREP-07: runs staged SharePoint deployment before app submission (fail deploy=%s)",
    async (failDeploy) => {
      const folder = mkdtempSync(join(tmpdir(), "fx-core-prepared-publish-"));
      const appPackage = createAppPackage();
      const sppkg = createSppkg();
      const requestOrder: string[] = [];
      let uploadedSharePointPackage: Buffer | undefined;
      let submittedPackage: Buffer | undefined;

      try {
        mkdirSync(join(folder, "env"));
        mkdirSync(join(folder, "sharepoint", "config"), { recursive: true });
        mkdirSync(join(folder, "appPackage", "build"), { recursive: true });
        writeFileSync(join(folder, "env", ".env.dev"), "TEAMSFX_ENV=dev\n");
        writeFileSync(join(folder, "sharepoint", "solution.sppkg"), sppkg);
        writeFileSync(join(folder, "appPackage", "build", "appPackage.dev.zip"), appPackage);
        writeFileSync(
          join(folder, "sharepoint", "config", "package-solution.json"),
          JSON.stringify({
            paths: { zippedPackage: "../solution.sppkg" },
            solution: { id: productId },
          })
        );
        writeFileSync(
          join(folder, "m365agents.yml"),
          [
            "version: v1.14",
            "environmentFolderPath: ./env",
            "publish:",
            "  - uses: spfx/preparePublication",
            "    with:",
            "      packagePath: ./sharepoint/solution.sppkg",
            "      appPackagePath: ./appPackage/build/appPackage.dev.zip",
            "    writeToEnvironmentFile:",
            "      sharePointPackagePath: STAGED_SPPKG",
            "      appPackagePath: STAGED_APP_PACKAGE",
            "  - uses: spfx/deploy",
            "    with:",
            "      packageSolutionPath: ./sharepoint/config/package-solution.json",
            "      createAppCatalogIfNotExist: true",
            "      packagePath: ${{STAGED_SPPKG}}",
            "  - uses: copilotAgent/publish",
            "    with:",
            "      appPackagePath: ${{STAGED_APP_PACKAGE}}",
            "    writeToEnvironmentFile:",
            "      titleId: M365_PUBLISHED_TITLE_ID",
            "      appId: M365_PUBLISHED_APP_ID",
            "",
          ].join("\n")
        );

        const respond = (config: Parameters<AxiosAdapter>[0], data: unknown) => ({
          config,
          data,
          headers: {},
          request: { method: config.method },
          status: 200,
          statusText: "OK",
        });
        axios.defaults.adapter = async (config) => {
          const url = `${config.baseURL ?? ""}${config.url ?? ""}`;
          if (url.includes("graph.microsoft.com") && url.includes("/sites/root")) {
            return respond(config, { webUrl: tenantUrl });
          }
          if (url.includes("/_api/SP_TenantSettings_Current")) {
            return respond(config, { CorporateCatalogUrl: `${tenantUrl}/appcatalog` });
          }
          if (url.includes("/Add(overwrite=true")) {
            requestOrder.push("sharepoint-upload");
            if (Buffer.isBuffer(config.data)) uploadedSharePointPackage = config.data;
            return respond(config, {});
          }
          if (url.includes("/tenantappcatalog/AvailableApps/GetById(")) {
            requestOrder.push("sharepoint-deploy");
            if (failDeploy) throw new Error("simulated SharePoint deployment failure");
            return respond(config, {});
          }
          if (url.includes("/config/v1/environment")) {
            return respond(config, { titlesServiceUrl: "https://titles.example.test" });
          }
          if (url.includes("/builder/v1/users/packages/status/")) {
            return respond(config, {
              titleId: "published-title-id",
              appId: "published-app-id",
            });
          }
          if (url.includes("/builder/v1/users/packages")) {
            requestOrder.push("app-submit");
            const data: unknown = config.data;
            if (
              typeof data === "object" &&
              data !== null &&
              "getBuffer" in data &&
              typeof data.getBuffer === "function"
            ) {
              submittedPackage = data.getBuffer();
            }
            return respond(config, { statusId: "upload-status" });
          }
          throw new Error(`Unexpected publish test request: ${url}`);
        };

        const tools = new MockTools();
        vi.spyOn(tools.tokenProvider.m365TokenProvider, "getJsonObject").mockResolvedValue(
          ok({ tid: tenantId, aud: "https://graph.microsoft.com" })
        );
        process.env.TEAMSFX_RAI_VALIDATION_ENABLED = "false";

        const result = await new FxCoreClient(tools).publish({
          platform: Platform.CLI,
          nonInteractive: true,
          projectPath: folder,
          env: "dev",
        });

        if (failDeploy) {
          expect(result.isErr()).toBe(true);
          expect(requestOrder).toEqual(["sharepoint-upload", "sharepoint-deploy"]);
          expect(submittedPackage).toBeUndefined();
        } else {
          expect(result.isErr() ? `${result.error.name}: ${result.error.message}` : undefined).toBe(
            undefined
          );
          if (result.isErr()) throw result.error;
          expect(requestOrder).toEqual(["sharepoint-upload", "sharepoint-deploy", "app-submit"]);
          expect(result.value.outputs["STAGED_SPPKG"]).toBeDefined();
          expect(result.value.outputs["STAGED_APP_PACKAGE"]).toBeDefined();
          expect(uploadedSharePointPackage).toEqual(
            readFileSync(join(folder, result.value.outputs["STAGED_SPPKG"]))
          );
          expect(submittedPackage).toBeDefined();
          expect(
            submittedPackage?.includes(
              readFileSync(join(folder, result.value.outputs["STAGED_APP_PACKAGE"]))
            )
          ).toBe(true);
          expect(result.value.outputs).toMatchObject({
            M365_PUBLISHED_TITLE_ID: "published-title-id",
            M365_PUBLISHED_APP_ID: "published-app-id",
          });
        }
      } finally {
        rmSync(folder, { recursive: true, force: true });
      }
    }
  );

  it("validate_RejectedPackage_ReturnsSuccessfulDomainOutcome", async () => {
    vi.spyOn(teamsappMgr, "validateTeamsAppForClient").mockResolvedValue(
      ok({
        status: "Rejected",
        errors: [
          {
            id: "invalid-manifest",
            content: "The manifest is invalid.",
            filePath: "manifest.json",
            shortCodeNumber: 1,
            title: "Invalid manifest",
            validationCategory: "manifest",
          },
        ],
        warnings: [],
        notes: [],
        addInDetails: {
          displayName: "Test app",
          developerName: "Test developer",
          version: "1.0.0",
          manifestVersion: "1.22",
        },
      })
    );
    const client = new FxCoreClient(new MockTools());

    const result = await client.validate({
      platform: Platform.CLI,
      projectPath: "project",
      "package-file": "app.zip",
    });

    expect(result.isOk()).toBe(true);
    if (result.isOk()) {
      expect(result.value).toEqual({
        valid: false,
        issues: [
          {
            severity: "error",
            message: "The manifest is invalid.",
            path: "manifest.json",
            code: "invalid-manifest",
            helpUrl: undefined,
          },
        ],
      });
    }
  });

  it("validate_WarningsAndNotes_MapsAllIssueSeverities", async () => {
    vi.spyOn(teamsappMgr, "validateTeamsAppForClient").mockResolvedValue(
      ok({
        status: "Accepted",
        errors: [],
        warnings: [
          {
            id: "warning-id",
            code: "warning-code",
            content: "",
            filePath: "",
            shortCodeNumber: 1,
            title: "Warning title",
            validationCategory: "manifest",
            helpUrl: "https://example.com/warning",
          },
        ],
        notes: [
          {
            id: "note-id",
            content: "Note content",
            shortCodeNumber: 2,
            title: "Note title",
            validationCategory: "manifest",
          },
        ],
        addInDetails: {
          displayName: "Test app",
          developerName: "Test developer",
          version: "1.0.0",
          manifestVersion: "1.22",
        },
      })
    );
    const client = new FxCoreClient(new MockTools());

    const result = await client.validate({
      platform: Platform.CLI,
      projectPath: "project",
      "package-file": "app.zip",
    });

    expect(result.isOk()).toBe(true);
    if (result.isOk()) {
      expect(result.value).toEqual({
        valid: true,
        issues: [
          {
            severity: "warning",
            message: "Warning title",
            path: undefined,
            code: "warning-code",
            helpUrl: "https://example.com/warning",
          },
          {
            severity: "info",
            message: "Note content",
            code: "note-id",
          },
        ],
      });
    }
  });

  it("getLaunchInfo_TitleId_ReturnsServiceDocument", async () => {
    const launchInfo = { acquisition: { titleId: "T_test" }, name: "Test agent" };
    const serviceSpy = vi
      .spyOn(PackageService.prototype, "getLaunchInfoByTitleId")
      .mockResolvedValue(launchInfo);
    const controller = new AbortController();
    const client = new FxCoreClient(new MockTools());

    const result = await client.getLaunchInfo({ titleId: "T_test" }, { signal: controller.signal });

    expect(result.isOk()).toBe(true);
    if (result.isOk()) {
      expect(result.value).toEqual(launchInfo);
    }
    expect(serviceSpy).toHaveBeenCalledWith("fakeToken", "T_test", controller.signal);
  });

  it("getLaunchInfo_ManifestId_ReturnsServiceDocument", async () => {
    const launchInfo = { acquisition: { titleId: "T_test" }, name: "Test agent" };
    const serviceSpy = vi
      .spyOn(PackageService.prototype, "getLaunchInfoByManifestId")
      .mockResolvedValue(launchInfo);
    const controller = new AbortController();
    const client = new FxCoreClient(new MockTools());

    const result = await client.getLaunchInfo(
      { manifestId: "manifest-id" },
      { signal: controller.signal }
    );

    expect(result.isOk()).toBe(true);
    if (result.isOk()) {
      expect(result.value).toEqual(launchInfo);
    }
    expect(serviceSpy).toHaveBeenCalledWith("fakeToken", "manifest-id", controller.signal);
  });

  it("getLaunchInfo_AbortedAfterRequest_ReturnsUserCancel", async () => {
    const controller = new AbortController();
    vi.spyOn(PackageService.prototype, "getLaunchInfoByTitleId").mockImplementation(async () => {
      controller.abort();
      return { name: "Test agent" };
    });
    const client = new FxCoreClient(new MockTools());

    const result = await client.getLaunchInfo({ titleId: "T_test" }, { signal: controller.signal });

    expect(result.isErr()).toBe(true);
    if (result.isErr()) {
      expect(result.error.name).toBe("UserCancel");
    }
  });

  it("getLaunchInfo_AbortedRequestThrows_ReturnsUserCancel", async () => {
    const controller = new AbortController();
    vi.spyOn(PackageService.prototype, "getLaunchInfoByTitleId").mockImplementation(async () => {
      controller.abort();
      throw new Error("aborted");
    });
    const client = new FxCoreClient(new MockTools());

    const result = await client.getLaunchInfo({ titleId: "T_test" }, { signal: controller.signal });

    expect(result.isErr()).toBe(true);
    if (result.isErr()) {
      expect(result.error.name).toBe("UserCancel");
    }
  });

  it("getLaunchInfo_RequestThrows_ReturnsAssembledError", async () => {
    vi.spyOn(PackageService.prototype, "getLaunchInfoByTitleId").mockRejectedValue(
      new Error("service failure")
    );
    const client = new FxCoreClient(new MockTools());

    const result = await client.getLaunchInfo({ titleId: "T_test" });

    expect(result.isErr()).toBe(true);
    if (result.isErr()) {
      expect(result.error.name).not.toBe("UserCancel");
    }
  });

  it("lifecycleOperations_EngineFailure_ReturnsOriginalError", async () => {
    const expectedError = new UserCancelError("test");
    vi.spyOn(FxCore.prototype, "provisionResources").mockResolvedValue(err(expectedError));
    vi.spyOn(FxCore.prototype, "packageTeamsAppCLIV3").mockResolvedValue(err(expectedError));
    vi.spyOn(FxCore.prototype, "publishApplication").mockResolvedValue(err(expectedError));
    vi.spyOn(teamsappMgr, "validateTeamsAppForClient").mockResolvedValue(err(expectedError));
    vi.spyOn(FxCore.prototype, "uninstall").mockResolvedValue(err(expectedError));
    const client = new FxCoreClient(new MockTools());
    const inputs = { platform: Platform.CLI, projectPath: "project", env: "dev" };

    const results = await Promise.all([
      client.provision(inputs),
      client.package(inputs),
      client.publish(inputs),
      client.validate(inputs),
      client.uninstall({ ...inputs, options: [] }),
    ]);

    for (const result of results) {
      expect(result.isErr()).toBe(true);
      if (result.isErr()) {
        expect(result.error).toBe(expectedError);
      }
    }
  });

  it("uninstall_CompletedOperation_ReturnsRemovedOptions", async () => {
    const uninstallSpy = vi.spyOn(FxCore.prototype, "uninstall").mockResolvedValue(ok());
    const controller = new AbortController();
    const client = new FxCoreClient(new MockTools());
    const inputs = {
      platform: Platform.CLI,
      projectPath: "project",
      options: ["m365-app"],
    };

    const result = await client.uninstall(inputs, { signal: controller.signal });

    expect(result.isOk()).toBe(true);
    if (result.isOk()) {
      expect(result.value.removed).toEqual(["m365-app"]);
    }
    expect(uninstallSpy).toHaveBeenCalledWith(
      expect.objectContaining({ abortSignal: controller.signal })
    );
  });
});
