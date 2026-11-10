// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { Platform } from "@microsoft/teamsfx-api";
import { describe, expect, it, vi } from "vitest";
import type { DriverContext } from "../../../../src/component/driver/interface/commonArgs";
import { MockTools } from "../../../core/utils";

vi.mock("../../../../src/actions", () => {
  throw new Error("Drivers must not depend on the SDK registry");
});
vi.mock("../../../../src/actions/contracts", () => {
  throw new Error("Driver contracts must not depend on the SDK registry");
});

describe("driver-owned action contracts", () => {
  it("ACTION-03: owns pure args contracts independently of the public registry", async () => {
    const { acquireArgs } = await import("../../../../src/component/driver/m365/args");
    const input = { appPackagePath: "app.zip", scope: "Personal" };
    const result = acquireArgs.validate(input);
    expect(result.isOk()).toBe(true);
    if (result.isOk()) expect(result.value).not.toBe(input);
    expect(acquireArgs.validate({ appPackagePath: 123 }).isErr()).toBe(true);
  });

  it("ACTION-03: validates driver inputs without importing the SDK registry", async () => {
    const { M365TitleAcquireDriver } =
      await import("../../../../src/component/driver/m365/acquire");
    const tools = new MockTools();
    const token = vi.spyOn(tools.tokenProvider.m365TokenProvider, "getAccessToken");
    const context: DriverContext = {
      azureAccountProvider: tools.tokenProvider.azureAccountProvider,
      m365TokenProvider: tools.tokenProvider.m365TokenProvider,
      logProvider: tools.logProvider,
      projectPath: process.cwd(),
      platform: Platform.CLI,
      ui: undefined,
      progressBar: undefined,
    };
    const result = await new M365TitleAcquireDriver().execute({ appPackagePath: "" }, context);
    expect(result.result.isErr()).toBe(true);
    if (result.result.isErr()) expect(result.result.error.name).toBe("InvalidActionInputError");
    expect(token).not.toHaveBeenCalled();
  });
});
