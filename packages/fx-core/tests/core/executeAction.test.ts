// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { Platform, UserError, err, ok } from "@microsoft/teamsfx-api";
import AdmZip from "adm-zip";
import axios from "axios";
import { mkdtemp, readdir, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import {
  M365TitleAcquireDriver,
  actionName as driverActionName,
} from "../../src/component/driver/m365/acquire";
import { ActionNames } from "../../src/actions";
import type { ExecutionResult } from "../../src/component/driver/interface/stepDriver";
import { PackageService } from "../../src/component/m365/packageService";
import { Lifecycle } from "../../src/component/configManager/lifecycle";
import { envUtil } from "../../src/component/utils/envUtil";
import { FxCoreClient } from "../../src/core/FxCoreClient";
import { FxCore } from "../../src/core/FxCore";
import { MockTools } from "./utils";
import { TelemetryEvent, TelemetryProperty } from "../../src/common/telemetry";
import { MosServiceScope } from "../../src/common/constants";
import { createContext } from "../../src/common/globalVars";
import { generateDriverContext } from "../../src/common/utils";
import { createDriverContext } from "../../src/component/driver/util/utils";
import { UserCancelError } from "../../src/error/common";

describe("standalone action execution", () => {
  let projectPath: string;
  let tools: MockTools;
  let client: FxCoreClient;

  beforeEach(async () => {
    projectPath = await mkdtemp(join(tmpdir(), "atk-actions-"));
    await writeFile(join(projectPath, "app.zip"), "fake package; outer service is stubbed");
    tools = new MockTools();
    client = new FxCoreClient(tools);
  });

  afterEach(async () => {
    vi.restoreAllMocks();
    vi.unstubAllEnvs();
    await rm(projectPath, { recursive: true, force: true });
  });

  it("ACTION-04: rejects invalid calls before driver execution or authentication", async () => {
    expect(client.executeAction).toBeTypeOf("function");
    const execute = vi.spyOn(M365TitleAcquireDriver.prototype, "execute");
    const token = vi.spyOn(tools.tokenProvider.m365TokenProvider, "getAccessToken");
    const calls: Array<[string, unknown, unknown]> = [
      ["script", {}, { projectPath }],
      ["teamsApp/extendToM365", {}, { projectPath }],
      ["teamsApp/extendToM365", { appPackagePath: 10 }, { projectPath }],
      ["teamsApp/extendToM365", { appPackagePath: "app.zip", extra: "secret" }, { projectPath }],
      ["teamsApp/extendToM365", { appPackagePath: "app.zip" }, { projectPath: "." }],
      ["teamsApp/extendToM365", { appPackagePath: "app.zip" }, undefined],
      [
        "teamsApp/extendToM365",
        { appPackagePath: "app.zip" },
        { projectPath, signal: new AbortController().signal },
      ],
    ];
    for (const args of calls) {
      const result: ExecutionResult = await Reflect.apply(client.executeAction, client, args);
      expect(result.result.isErr()).toBe(true);
    }
    expect(execute).not.toHaveBeenCalled();
    expect(token).not.toHaveBeenCalled();
  });

  it("ACTION-05: installs once using a snapshot and returns outputs without lifecycle side effects", async () => {
    expect(client.executeAction).toBeTypeOf("function");
    const service = vi
      .spyOn(PackageService.prototype, "sideLoading")
      .mockResolvedValue(["U_title", "app-id", "https://example.test/share"]);
    const token = vi.spyOn(tools.tokenProvider.m365TokenProvider, "getAccessToken");
    const driver = vi.spyOn(M365TitleAcquireDriver.prototype, "execute");
    const lifecycle = vi.spyOn(Lifecycle.prototype, "execute");
    const readEnv = vi.spyOn(envUtil, "readEnv");
    const writeEnv = vi.spyOn(envUtil, "writeEnv");
    const environment = { ...process.env };
    const args = { appPackagePath: "app.zip" };
    expect(ActionNames.Acquire).toBe(driverActionName);
    const pending = client.executeAction(ActionNames.Acquire, args, { projectPath });
    args.appPackagePath = "changed.zip";
    const result = await pending;

    expect(result.result.isOk()).toBe(true);
    if (result.result.isOk()) {
      expect(Object.fromEntries(result.result.value)).toEqual({
        titleId: "U_title",
        appId: "app-id",
        shareLink: "https://example.test/share",
      });
    }
    expect(service).toHaveBeenCalledExactlyOnceWith(
      "fakeToken",
      join(projectPath, "app.zip"),
      undefined
    );
    expect(token).toHaveBeenCalledWith({ scopes: MosServiceScope() });
    expect(driver).toHaveBeenCalledWith(
      { appPackagePath: "app.zip" },
      expect.objectContaining({
        projectPath,
        platform: Platform.CLI,
        nonInteractive: true,
        m365TokenProvider: tools.tokenProvider.m365TokenProvider,
        ui: undefined,
      })
    );
    expect(lifecycle).not.toHaveBeenCalled();
    expect(readEnv).not.toHaveBeenCalled();
    expect(writeEnv).not.toHaveBeenCalled();
    expect(process.env).toEqual(environment);
    expect(await readdir(projectPath)).toEqual(["app.zip"]);
  });

  it("ACTION-06: keeps explicit installation output mappings and empty mapped share links", async () => {
    vi.spyOn(PackageService.prototype, "sideLoading").mockResolvedValue(["U_title", "app-id", ""]);
    const context = {
      projectPath,
      platform: Platform.CLI,
      nonInteractive: true,
      azureAccountProvider: tools.tokenProvider.azureAccountProvider,
      m365TokenProvider: tools.tokenProvider.m365TokenProvider,
      logProvider: tools.logProvider,
      ui: undefined,
      progressBar: undefined,
    };
    const driver = new M365TitleAcquireDriver();
    const result = await driver.execute(
      { appPackagePath: "app.zip" },
      context,
      new Map([
        ["titleId", "TITLE"],
        ["appId", "APP"],
        ["shareLink", "LINK"],
      ])
    );
    expect(result.result.isOk()).toBe(true);
    if (result.result.isOk()) {
      expect(Object.fromEntries(result.result.value)).toEqual({
        TITLE: "U_title",
        APP: "app-id",
        LINK: "",
      });
    }
    const invalid = await driver.execute(
      { appPackagePath: "app.zip" },
      context,
      new Map([["titleId", "TITLE"]])
    );
    expect(invalid.result.isErr()).toBe(true);
  });

  it("ACTION-07: rejects uninstall as an action and preserves its existing API route", async () => {
    const legacy = vi.spyOn(FxCore.prototype, "uninstall").mockResolvedValue(ok(undefined));
    const service = vi.spyOn(PackageService.prototype, "unacquire").mockResolvedValue();
    const token = vi.spyOn(tools.tokenProvider.m365TokenProvider, "getAccessToken");
    const result: ExecutionResult = await Reflect.apply(client.executeAction, client, [
      "teamsApp/uninstallFromM365",
      { titleId: "U_title" },
      { projectPath },
    ]);
    expect(result.result.isErr()).toBe(true);
    if (result.result.isErr()) expect(result.result.error.name).toBe("UnsupportedAction");
    expect(legacy).not.toHaveBeenCalled();
    expect(service).not.toHaveBeenCalled();
    expect(token).not.toHaveBeenCalled();
    const inputs = { platform: Platform.CLI, projectPath, options: ["m365-app"] };
    const existing = await client.uninstall(inputs);
    expect(existing.isOk()).toBe(true);
    expect(legacy).toHaveBeenCalledExactlyOnceWith(expect.objectContaining(inputs));
  });

  it.each([false, true])(
    "ACTION-08: rejects unsupported acquisition cancellation (aborted=%s)",
    async (aborted) => {
      const controller = new AbortController();
      if (aborted) controller.abort();
      const driver = vi.spyOn(M365TitleAcquireDriver.prototype, "execute");
      const token = vi.spyOn(tools.tokenProvider.m365TokenProvider, "getAccessToken");
      const result = await client.executeAction(
        "teamsApp/extendToM365",
        { appPackagePath: "app.zip" },
        { projectPath, signal: controller.signal }
      );
      expect(result.result.isErr()).toBe(true);
      if (result.result.isErr()) expect(result.result.error.name).toBe("UnsupportedCancellation");
      expect(driver).not.toHaveBeenCalled();
      expect(token).not.toHaveBeenCalled();
    }
  );

  it("ACTION-09: isolates instance providers and propagates authentication failures", async () => {
    expect(client.executeAction).toBeTypeOf("function");
    const secondTools = new MockTools();
    const second = new FxCore(secondTools);
    const firstToken = vi
      .spyOn(tools.tokenProvider.m365TokenProvider, "getAccessToken")
      .mockResolvedValue(ok("first"));
    const secondToken = vi
      .spyOn(secondTools.tokenProvider.m365TokenProvider, "getAccessToken")
      .mockResolvedValue(ok("second"));
    const install = vi
      .spyOn(PackageService.prototype, "sideLoading")
      .mockResolvedValue(["U_title", "app-id", ""]);
    await Promise.all([
      client.executeAction("teamsApp/extendToM365", { appPackagePath: "app.zip" }, { projectPath }),
      second.executeAction("teamsApp/extendToM365", { appPackagePath: "app.zip" }, { projectPath }),
    ]);
    expect(install).toHaveBeenCalledWith("first", join(projectPath, "app.zip"), undefined);
    expect(install).toHaveBeenCalledWith("second", join(projectPath, "app.zip"), undefined);
    expect(firstToken).toHaveBeenCalledWith({ scopes: MosServiceScope() });
    expect(secondToken).toHaveBeenCalledWith({ scopes: MosServiceScope() });
    const failure = new UserError({
      source: "test",
      name: "InteractionRequired",
      message: "authentication unavailable",
    });

    firstToken.mockResolvedValue(err(failure));
    install.mockClear();
    const result = await client.executeAction(
      "teamsApp/extendToM365",
      { appPackagePath: "app.zip" },
      { projectPath }
    );
    expect(result.result.isErr()).toBe(true);
    if (result.result.isErr()) expect(result.result.error).toBe(failure);
    expect(install).not.toHaveBeenCalled();
  });

  it("ACTION-11: permits provider-owned login while keeping action parameters noninteractive", async () => {
    const provider = tools.tokenProvider.m365TokenProvider;
    const token = vi.spyOn(provider, "getAccessToken").mockImplementation(async (request) =>
      request.showDialog === false
        ? err(
            new UserError({
              source: "auth",
              name: "InteractionRequired",
              message: "login required",
            })
          )
        : ok("interactive-login-token")
    );
    const install = vi
      .spyOn(PackageService.prototype, "sideLoading")
      .mockResolvedValue(["U_title", "app-id", ""]);
    const driver = vi.spyOn(M365TitleAcquireDriver.prototype, "execute");

    const result = await client.executeAction(
      ActionNames.Acquire,
      { appPackagePath: "app.zip" },
      { projectPath }
    );

    expect(result.result.isOk()).toBe(true);
    expect(token).toHaveBeenCalledExactlyOnceWith({ scopes: MosServiceScope() });
    expect(install).toHaveBeenCalledExactlyOnceWith(
      "interactive-login-token",
      join(projectPath, "app.zip"),
      undefined
    );
    const context = driver.mock.calls[0][1];
    expect(context.m365TokenProvider).toBe(provider);
    expect(context.nonInteractive).toBe(true);
    expect(context.ui).toBeUndefined();
  });

  it("ACTION-11: provider login cancellation propagates before any service write", async () => {
    const cancellation = new UserCancelError();
    vi.spyOn(tools.tokenProvider.m365TokenProvider, "getAccessToken").mockResolvedValue(
      err(cancellation)
    );
    const install = vi.spyOn(PackageService.prototype, "sideLoading");

    const result = await client.executeAction(
      ActionNames.Acquire,
      { appPackagePath: "app.zip" },
      { projectPath }
    );

    expect(result.result.isErr()).toBe(true);
    if (result.result.isErr()) expect(result.result.error).toBe(cancellation);
    expect(install).not.toHaveBeenCalled();
  });

  it("ACTION-06: keeps existing shared context shapes unchanged", () => {
    const expServiceProvider = { getTreatmentVariableAsync: vi.fn(async () => undefined) };
    new FxCore({ ...tools, expServiceProvider });
    const inputs = { projectPath, platform: Platform.CLI };
    const context = createContext();
    expect(context).not.toHaveProperty("expServiceProvider");
    expect(generateDriverContext(context, inputs)).not.toHaveProperty("expServiceProvider");
    expect(createDriverContext(inputs)).not.toHaveProperty("expServiceProvider");
  });

  it("ACTION-04: converts unexpected driver exceptions into structured errors", async () => {
    vi.spyOn(M365TitleAcquireDriver.prototype, "execute").mockRejectedValue(
      new Error("driver failed")
    );
    const result = await client.executeAction(
      "teamsApp/extendToM365",
      { appPackagePath: "app.zip" },
      { projectPath }
    );
    expect(result.result.isErr()).toBe(true);
    if (result.result.isErr()) expect(result.result.error.message).toContain("driver failed");
  });

  it("ACTION-05: returns the native driver result object and domain fields unchanged", async () => {
    const native = {
      result: ok(new Map<string, string>()),
      summaries: ["completed"],
      report: { operationId: "test-operation", outcome: "published" },
    };
    vi.spyOn(M365TitleAcquireDriver.prototype, "execute").mockResolvedValue(native);
    const result = await client.executeAction(
      ActionNames.Acquire,
      { appPackagePath: "app.zip" },
      { projectPath }
    );
    expect(result).toBe(native);
  });

  it("ACTION-09: reuses existing service configuration when installing a real ZIP through fake HTTP", async () => {
    vi.stubEnv("TEAMSFX_RAI_VALIDATION_ENABLED", undefined);
    vi.stubEnv("TEAMSFX_INCREMENTAL_SKILL_UPLOAD", "false");
    const zip = new AdmZip();
    zip.addFile(
      "manifest.json",
      Buffer.from(JSON.stringify({ manifestVersion: "1.24", id: "app-id" }))
    );
    await writeFile(join(projectPath, "app.zip"), zip.toBuffer());
    const ownExperiments = { getTreatmentVariableAsync: vi.fn(async () => undefined) };
    const ownTelemetry = vi.spyOn(tools.telemetryReporter, "sendTelemetryEvent");
    const ownClient = new FxCoreClient({ ...tools, expServiceProvider: ownExperiments });
    const http = axios.create({
      adapter: async (config) => {
        let data: unknown;
        if (config.url === "/config/v1/environment") {
          data = { titlesServiceUrl: "https://service.test" };
        } else if (config.url === "/builder/v1/users/packages") {
          data = { statusId: "test-status" };
        } else if (config.url === "/builder/v1/users/packages/status/test-status") {
          data = { titleId: "U_title", appId: "app-id" };
        } else {
          throw new Error(`Unexpected HTTP request: ${config.url}`);
        }
        return {
          data,
          status: 200,
          statusText: "OK",
          headers: {},
          config,
          request: { method: config.method },
        };
      },
    });
    vi.spyOn(axios, "create").mockReturnValue(http);
    const result = await ownClient.executeAction(
      "teamsApp/extendToM365",
      { appPackagePath: "app.zip" },
      { projectPath }
    );
    expect(result.result.isOk()).toBe(true);
    expect(ownExperiments.getTreatmentVariableAsync).toHaveBeenCalled();
    expect(ownTelemetry).toHaveBeenCalledWith(
      TelemetryEvent.MosSideloadEnd,
      expect.objectContaining({ [TelemetryProperty.MosTitleId]: "U_title" }),
      undefined
    );
  });
});
