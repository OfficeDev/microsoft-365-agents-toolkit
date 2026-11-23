// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { exec } from "child_process";
import { mkdtemp, rm } from "fs/promises";
import { tmpdir } from "os";
import path from "path";
import * as vscode from "vscode";
import { ok } from "@microsoft/teamsfx-api";
import { envUtil } from "@microsoft/teamsfx-core";
import { expect, vi } from "vitest";
import { LaunchDesktopClientTerminal } from "../../src/debug/taskTerminal/launchDesktopClientTerminal";
import { ExtensionErrors } from "../../src/error/error";
import * as globalVariables from "../../src/globalVariables";
import { mockValue } from "../mocks/vitestMockUtils";

vi.mock("child_process", async (importOriginal) => {
  const childProcess = await importOriginal<typeof import("child_process")>();
  return {
    ...childProcess,
    exec: vi.fn(() => {
      const child = new childProcess.ChildProcess();
      queueMicrotask(() => child.emit("close", 0));
      return child;
    }),
  };
});

describe("LaunchDesktopClientTerminal", () => {
  const appId = "00000000-0000-4000-8000-000000000001";
  const url = `teams.microsoft.com/l/app/${appId}?installAppPackage=true`;
  let projectPath: string;

  function launch(args?: { url: unknown }) {
    return new LaunchDesktopClientTerminal({ type: "teamsfx", args }).do();
  }

  beforeEach(async () => {
    projectPath = await mkdtemp(path.join(tmpdir(), "atk-desktop-launch-"));
    mockValue(globalVariables, "workspaceUri", vscode.Uri.file(projectPath));
    mockValue(
      globalVariables.tools.tokenProvider.m365TokenProvider,
      "getStatus",
      vi.fn().mockResolvedValue(ok({ status: "SignedOut" }))
    );
    mockValue(vscode.workspace, "isTrusted", true);
    vi.spyOn(vscode.env, "openExternal").mockResolvedValue(true);
    vi.mocked(exec).mockClear();
  });

  afterEach(async () => {
    vi.restoreAllMocks();
    await rm(projectPath, { recursive: true, force: true });
  });

  it.each([
    { platform: "win32", scheme: "msteams" },
    { platform: "darwin", scheme: "msteams" },
    { platform: "linux", scheme: "https" },
  ])(
    "AC-01: opens the resolved $scheme link on $platform without a shell",
    async ({ platform, scheme }) => {
      mockValue(process, "platform", platform);
      vi.spyOn(envUtil, "readEnv").mockResolvedValue(ok({ TEAMS_APP_ID: appId }));

      const result = await launch({
        url: "teams.microsoft.com/l/app/${{local:TEAMS_APP_ID}}?installAppPackage=true",
      });

      expect(result.isOk()).toBe(true);
      expect(exec).not.toHaveBeenCalled();
      expect(vscode.env.openExternal).toHaveBeenCalledExactlyOnceWith(
        expect.objectContaining({
          scheme,
          authority: "teams.microsoft.com",
          path: `/l/app/${appId}`,
          query: "installAppPackage=true",
        })
      );
    }
  );

  it.each([
    { label: "missing args", args: undefined },
    { label: "non-string URL", args: { url: 123 } },
    { label: "explicit scheme", args: { url: `https://${url}` } },
    { label: "unexpected host", args: { url: "example.invalid/l/app/example" } },
    { label: "user information", args: { url: `user@${url}` } },
    { label: "unexpected port", args: { url: "teams.microsoft.com:8443/l/app/example" } },
  ])("AC-02: rejects $label without opening anything", async ({ args }) => {
    const result = await launch(args);

    expect(result._unsafeUnwrapErr().name).toBe(ExtensionErrors.TaskDefinitionError);
    expect(vscode.env.openExternal).not.toHaveBeenCalled();
    expect(exec).not.toHaveBeenCalled();
  });

  it("AC-02: rejects an unsupported host supplied through an environment variable", async () => {
    vi.spyOn(envUtil, "readEnv").mockResolvedValue(ok({ TEAMS_HOST: "example.invalid" }));

    const result = await launch({ url: "${{local:TEAMS_HOST}}/l/app/example" });

    expect(result._unsafeUnwrapErr().name).toBe(ExtensionErrors.TaskDefinitionError);
    expect(vscode.env.openExternal).not.toHaveBeenCalled();
    expect(exec).not.toHaveBeenCalled();
  });

  it("AC-03: refuses to launch from an untrusted workspace before checking the account", async () => {
    mockValue(vscode.workspace, "isTrusted", false);

    const result = await launch({ url });

    expect(result.isErr()).toBe(true);
    expect(globalVariables.tools.tokenProvider.m365TokenProvider.getStatus).not.toHaveBeenCalled();
    expect(vscode.env.openExternal).not.toHaveBeenCalled();
    expect(exec).not.toHaveBeenCalled();
  });

  it.each([false, new Error("Unable to open link")])(
    "AC-04: reports an external opener failure (%s) without a shell fallback",
    async (failure) => {
      mockValue(process, "platform", "win32");
      if (failure instanceof Error) {
        vi.mocked(vscode.env.openExternal).mockRejectedValue(failure);
      } else {
        vi.mocked(vscode.env.openExternal).mockResolvedValue(failure);
      }

      const result = await launch({ url });

      expect(result._unsafeUnwrapErr().name).toBe(ExtensionErrors.LaunchTeamsDesktopClientError);
      expect(exec).not.toHaveBeenCalled();
    }
  );
});
