// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import * as cp from "child_process";
import crossSpawn from "cross-spawn";
import fs from "fs-extra";
import { tmpdir } from "os";
import path from "path";
import { PassThrough } from "stream";
import * as vscode from "vscode";
import { expect, vi } from "vitest";
import { OfficeDevTerminal, TriggerCmdType } from "../../src/debug/taskTerminal/officeDevTerminal";
import * as globalVariables from "../../src/globalVariables";
import { mockValue } from "../mocks/vitestMockUtils";

vi.mock("child_process", async (importOriginal) => {
  const actual = await importOriginal<typeof import("child_process")>();
  return { ...actual, spawn: vi.fn() };
});

const manifestCommands = [
  {
    trigger: TriggerCmdType.triggerValidate,
    args: ["--yes", "office-addin-manifest", "validate"],
    suffix: [],
  },
  {
    trigger: TriggerCmdType.triggerStopDebug,
    args: ["office-addin-debugging", "stop"],
    suffix: [],
  },
  {
    trigger: TriggerCmdType.triggerGenerateGUID,
    args: ["--yes", "office-addin-manifest", "modify"],
    suffix: ["--guid"],
  },
];

describe("OfficeDevTerminal process boundaries", () => {
  let workspace: string;

  beforeEach(async () => {
    workspace = await fs.mkdtemp(path.join(tmpdir(), "atk-office-spawn-"));
    const workspaceUri = vscode.Uri.file(workspace);
    workspace = workspaceUri.fsPath;
    mockValue(globalVariables, "workspaceUri", workspaceUri);
  });

  afterEach(async () => {
    vi.restoreAllMocks();
    vi.unstubAllEnvs();
    await fs.remove(workspace);
  });

  for (const { trigger, args, suffix } of manifestCommands) {
    it(`requests safe spawning for ${trigger} with a literal manifest filename`, async () => {
      const manifest = "manifest&node marker.cjs&.xml";
      await fs.writeFile(path.join(workspace, manifest), "<OfficeApp />");
      const child = new cp.ChildProcess();
      child.stdout = new PassThrough();
      child.stderr = new PassThrough();
      // Prevent the old implementation from launching npx while checking the regression.
      vi.spyOn(cp, "spawn").mockReturnValue(child);
      const spawn = vi.spyOn(crossSpawn, "spawn").mockReturnValue(child);
      const terminal = new OfficeDevTerminal();
      const output: string[] = [];
      terminal.onDidWrite((data) => output.push(data));

      terminal.handleInput(trigger);

      expect(spawn).toHaveBeenCalledWith("npx", [...args, manifest, ...suffix], {
        cwd: workspace,
        shell: false,
        detached: false,
      });
      child.stdout.emit("data", "stdout\n");
      child.stderr.emit("data", "stderr\n");
      child.emit("exit", 0);
      expect(output.join("")).toContain("stdout\r\n");
      expect(output.join("")).toContain("stderr\r\n");
    });
  }

  it("requests safe spawning for npm install", () => {
    const child = new cp.ChildProcess();
    vi.spyOn(cp, "spawn").mockReturnValue(child);
    const spawn = vi.spyOn(crossSpawn, "spawn").mockReturnValue(child);

    new OfficeDevTerminal().handleInput(TriggerCmdType.triggerInstall);

    expect(spawn).toHaveBeenCalledWith("npm", ["install", "--color=always"], {
      cwd: workspace,
      shell: false,
      detached: false,
    });
  });

  it("reports process launch failures and closes the terminal with an error", () => {
    const child = new cp.ChildProcess();
    vi.spyOn(crossSpawn, "spawn").mockReturnValue(child);
    const terminal = new OfficeDevTerminal();
    const output: string[] = [];
    const closeCodes: number[] = [];
    terminal.onDidWrite((data) => output.push(data));
    terminal.onDidClose?.((code) => closeCodes.push(code));

    terminal.handleInput(TriggerCmdType.triggerInstall);
    child.emit("error", new Error("npm executable not found"));

    expect(output.join("")).toContain("npm executable not found");
    expect(closeCodes).toEqual([1]);
  });

  for (const { trigger, args, suffix } of manifestCommands) {
    for (const shimDirectory of [path.join("node_modules", ".bin"), "global bin"]) {
      for (const manifest of [
        "manifest with spaces.xml",
        "manifest&node marker.cjs&.xml",
        "manifest%ATK_TEST_EXPANSION%.xml",
        "manifest^caret!.xml",
        "manifest (copy).xml",
        "manifest'$(printf literal);.xml",
        "manifest$HOME`echo`.xml",
      ]) {
        it(`preserves ${JSON.stringify(manifest)} through ${shimDirectory} for ${trigger}`, async () => {
          const bin = path.join(workspace, shimDirectory);
          await fs.ensureDir(bin);
          const capture = path.join(bin, "capture.cjs");
          await fs.writeFile(
            capture,
            "process.stdout.write(JSON.stringify(process.argv.slice(2)));"
          );
          const shim = path.join(bin, process.platform === "win32" ? "npx.cmd" : "npx");
          await fs.writeFile(
            shim,
            process.platform === "win32"
              ? `@"${process.execPath}" "%~dp0capture.cjs" %*\r\n`
              : `#!/usr/bin/env node\nrequire("./capture.cjs");\n`,
            { mode: 0o755 }
          );
          await fs.writeFile(path.join(workspace, manifest), "<OfficeApp />");
          await fs.writeFile(
            path.join(workspace, "marker.cjs"),
            'require("fs").writeFileSync("injected.txt", "unexpected");'
          );
          vi.stubEnv("ATK_TEST_EXPANSION", "expanded");
          const spawnProcess = crossSpawn.spawn;
          const children: cp.ChildProcess[] = [];
          vi.spyOn(crossSpawn, "spawn").mockImplementation((_command, commandArgs, options) => {
            const child = spawnProcess(shim, commandArgs, options);
            children.push(child);
            return child;
          });
          const terminal = new OfficeDevTerminal();
          let output = "";
          terminal.onDidWrite((data) => {
            output += data;
          });

          terminal.handleInput(trigger);

          expect(children).toHaveLength(1);
          const child = children[0];
          const exitCode = await new Promise<number | null>((resolve, reject) => {
            child.once("error", reject);
            child.once("close", resolve);
          });
          expect(exitCode, output).toBe(0);
          expect(output).toContain(JSON.stringify([...args, manifest, ...suffix]));
          expect(await fs.pathExists(path.join(workspace, "injected.txt"))).toBe(false);
        });
      }
    }
  }
});
