// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { CLIContext, UserError, err, ok } from "@microsoft/teamsfx-api";
import {
  PackageService,
  PolicyDestinationResult,
  PolicyPublicationReport,
  workIqPolicyClient,
} from "@microsoft/teamsfx-core";
import { ActionNames, getActionDefinition } from "@microsoft/teamsfx-core/actions";
import { mkdtemp, readdir, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { getFxCore } from "../../../src/activate";
import { engine } from "../../../src/commands/engine";
import { helper } from "../../../src/commands/helper";
import { rootCommand } from "../../../src/commands/models/root";
import { logger } from "../../../src/commonlib/logger";
import M365TokenProvider from "../../../src/commonlib/M365TokenProviderWrapper";

describe("hidden action command", () => {
  let projectPath: string;
  const core = getFxCore();
  const report: PolicyPublicationReport = {
    reportVersion: 1,
    operationId: "operation",
    outcome: "published",
    nextAction: "none",
    enforcement: "not-verified",
    mos: { state: "confirmed" },
    workIq: { state: "confirmed" },
  };
  const policyArgs = [
    ActionNames.PublishPolicy,
    "--title-id",
    "U_title",
    "--policy-path",
    "policy.json",
    "--caller-id",
    "11111111-2222-4333-8444-555555555555",
  ];

  beforeEach(async () => {
    vi.stubEnv("TEAMSFX_CLI_BIN_NAME", "atk");
    vi.stubEnv("CI_ENABLED", undefined);
    projectPath = await mkdtemp(path.join(tmpdir(), "atk-cli-action-"));
    await writeFile(path.join(projectPath, "policy.json"), '{"rules":[]}');
    vi.spyOn(core, "checkProjectType").mockResolvedValue(
      err(new UserError({ source: "test", name: "NotProject", message: "Ordinary directory" }))
    );
    vi.spyOn(core, "projectVersionCheck");
    vi.spyOn(core, "executeAction").mockResolvedValue({
      result: ok(new Map()),
      summaries: [],
    });
    vi.spyOn(logger, "info").mockImplementation(() => {});
    vi.spyOn(logger, "debug").mockImplementation(() => {});
    vi.spyOn(logger, "error").mockImplementation(() => {});
    vi.spyOn(M365TokenProvider, "getAccessToken");
  });

  afterEach(async () => {
    vi.restoreAllMocks();
    vi.unstubAllEnvs();
    await rm(projectPath, { recursive: true, force: true });
  });

  function contextFor(args: string[]) {
    const found = engine.findCommand(rootCommand, ["execute", "action", ...args]);
    expect(found.cmd.fullName).toMatch(/^atk execute action/);
    const context: CLIContext = {
      command: found.cmd,
      optionValues: {},
      globalOptionValues: { telemetry: false },
      argumentValues: [],
      telemetryProperties: {},
    };
    return { context, remainingArgs: found.remainingArgs };
  }

  async function run(args: string[]) {
    const { context, remainingArgs } = contextFor(args);
    const argv = process.argv;
    process.argv = [process.execPath, "atk", "execute", "action", ...args];
    try {
      const result = await engine.execute(context, rootCommand, remainingArgs);
      return { result, context };
    } finally {
      process.argv = argv;
    }
  }

  function printedResults(): unknown[] {
    return vi
      .mocked(logger.info)
      .mock.calls.flatMap(([message]) =>
        typeof message === "string" && message.startsWith("{") ? [JSON.parse(message)] : []
      );
  }

  it("CLI-ACTION-01: hides the entry while preserving group and schema-derived leaf help", async () => {
    const execute = rootCommand.commands?.find((command) => command.name === "execute");
    expect(execute?.hidden).toBe(true);
    expect(helper.formatHelp(rootCommand)).not.toMatch(/^\s+execute(?:\s|$)/m);
    const group = await run(["--help"]);
    expect(group.result.isOk()).toBe(true);
    for (const name of Object.values(ActionNames)) {
      expect(vi.mocked(logger.info).mock.calls.flat().join("\n")).toContain(name);
      const leaf = await run([name, "--help"]);
      expect(leaf.result.isOk()).toBe(true);
      const definition = getActionDefinition(name);
      expect(definition.isOk()).toBe(true);
      if (definition.isOk()) {
        expect(leaf.context.command.description).toContain(definition.value.argsSchema.description);
      }
    }
    const leaf = contextFor([ActionNames.PublishPolicy, "--help"]).context.command;
    expect(helper.formatHelp(leaf, rootCommand)).not.toContain("--folder");
    expect(leaf.description).toContain("current working directory");
    expect(leaf.options).toEqual(
      expect.arrayContaining([
        expect.objectContaining({ name: "title-id", questionName: "titleId", required: true }),
        expect.objectContaining({
          name: "policy-path",
          questionName: "policyPath",
          required: true,
        }),
        expect.objectContaining({ name: "caller-id", questionName: "callerId", required: true }),
      ])
    );
    expect(leaf.description).toMatch(/writ/i);
    const acquire = contextFor([ActionNames.Acquire]).context.command;
    expect(helper.formatHelp(acquire, rootCommand)).not.toContain("--folder");
    expect(acquire.options).toContainEqual(
      expect.objectContaining({
        name: "scope",
        required: false,
        choices: ["Personal", "Shared", "Tenant", "personal", "shared", "tenant"],
      })
    );
    expect(core.executeAction).not.toHaveBeenCalled();
    expect(M365TokenProvider.getAccessToken).not.toHaveBeenCalled();
  });

  it("CLI-ACTION-01: hides inherited interactive options only for the execute command tree", () => {
    expect(helper.formatHelp(rootCommand)).toContain("--interactive");
    const install = engine.findCommand(rootCommand, ["install"]).cmd;
    expect(helper.formatHelp(install, rootCommand)).toContain("--interactive");
    for (const args of [
      ["execute"],
      ["execute", "action"],
      ...Object.values(ActionNames).map((name) => ["execute", "action", name]),
    ]) {
      const command = engine.findCommand(rootCommand, args).cmd;
      const help = helper.formatHelp(command, rootCommand);
      expect(help).not.toContain("--interactive");
      expect(help).not.toMatch(/(?:^|\s)-i(?:\s|$)/);
    }
  });

  it.each([
    ["execute"],
    ["execute", "action"],
    ["execute", "action", ActionNames.PublishPolicy],
    ["execute", "action", ActionNames.Acquire],
  ])("CLI-ACTION-04: help separates argument mode from provider-owned sign-in (%j)", (...args) => {
    const command = engine.findCommand(rootCommand, args).cmd;
    expect(command.description).toContain("Action arguments must be supplied explicitly.");
    expect(command.description).toContain("configured token provider");
    const help = helper.formatHelp(command, rootCommand);
    expect(help).toMatch(/sign-in\s+prompts/);
    expect(help).not.toContain("--interactive");
  });

  it("CLI-ACTION-02: resolves policy paths from CWD and derives their parent independently of folder", async () => {
    const policyPath = path.join(projectPath, "policy.json");
    vi.spyOn(process, "cwd").mockReturnValue(path.dirname(projectPath));
    const relativePolicyPath = path.relative(process.cwd(), policyPath);
    expect(path.isAbsolute(relativePolicyPath)).toBe(false);
    const { result, context } = await run([
      ...policyArgs,
      "--policy-path",
      relativePolicyPath,
      "--folder",
      path.dirname(projectPath),
    ]);
    expect(result.isOk()).toBe(true);
    expect(core.executeAction).toHaveBeenCalledExactlyOnceWith(
      ActionNames.PublishPolicy,
      {
        titleId: "U_title",
        policyPath,
        callerId: "11111111-2222-4333-8444-555555555555",
      },
      { projectPath }
    );
    expect(context.globalOptionValues.interactive).toBe(false);
    expect(core.projectVersionCheck).not.toHaveBeenCalled();
  });

  it("CLI-ACTION-02: accepts absolute policy paths outside CWD without requiring folder", async () => {
    const policyPath = path.join(projectPath, "policy.json");
    const { result } = await run([...policyArgs, "--policy-path", policyPath]);
    expect(result.isOk()).toBe(true);
    expect(core.executeAction).toHaveBeenCalledExactlyOnceWith(
      ActionNames.PublishPolicy,
      {
        titleId: "U_title",
        policyPath,
        callerId: "11111111-2222-4333-8444-555555555555",
      },
      { projectPath }
    );
  });

  it("CLI-ACTION-02: preserves explicit hidden folder semantics for other actions", async () => {
    const { result } = await run([
      ActionNames.Acquire,
      "--app-package-path",
      "app.zip",
      "--folder",
      projectPath,
    ]);
    expect(result.isOk()).toBe(true);
    expect(core.executeAction).toHaveBeenCalledExactlyOnceWith(
      ActionNames.Acquire,
      { appPackagePath: "app.zip" },
      { projectPath }
    );
  });

  it("CLI-ACTION-02: defaults folder to cwd and does not inject omitted action defaults", async () => {
    const { result } = await run([ActionNames.Acquire, "--app-package-path", "app.zip"]);
    expect(result.isOk()).toBe(true);
    expect(core.executeAction).toHaveBeenCalledExactlyOnceWith(
      ActionNames.Acquire,
      { appPackagePath: "app.zip" },
      { projectPath: process.cwd() }
    );
  });

  it("CLI-ACTION-02: preserves valid inline string and enum values", async () => {
    const { result } = await run([
      ActionNames.Acquire,
      "--app-package-path=app.zip",
      "--scope=Shared",
    ]);
    expect(result.isOk()).toBe(true);
    expect(core.executeAction).toHaveBeenCalledExactlyOnceWith(
      ActionNames.Acquire,
      { appPackagePath: "app.zip", scope: "Shared" },
      { projectPath: process.cwd() }
    );
  });

  it("CLI-ACTION-02: preserves equals characters in separate string values", async () => {
    const { result } = await run([ActionNames.Acquire, "--app-package-path", "app=v1.zip"]);
    expect(result.isOk()).toBe(true);
    expect(core.executeAction).toHaveBeenCalledExactlyOnceWith(
      ActionNames.Acquire,
      { appPackagePath: "app=v1.zip" },
      { projectPath: process.cwd() }
    );
  });

  it.each([
    [ActionNames.PublishPolicy],
    [ActionNames.PublishPolicy, "--title-id", "U_title", "--policy-path", "policy.json"],
    [...policyArgs.slice(0, -1), "   "],
    [...policyArgs, "--policy-path", "   "],
    [...policyArgs, "--policy-snapshot", "secret-value"],
    [ActionNames.Acquire, "--app-package-path", "app.zip", "--scope", "invalid"],
    [ActionNames.Acquire, "--app-package-path", ""],
    ["arbitrary/module"],
  ])("CLI-ACTION-03: rejects invalid input before executing (%j)", async (...args) => {
    const { result } = await run(args);
    expect(result.isErr()).toBe(true);
    expect(core.executeAction).not.toHaveBeenCalled();
    expect(M365TokenProvider.getAccessToken).not.toHaveBeenCalled();
  });

  it("CLI-ACTION-04: rejects explicit interactive mode without executing", async () => {
    const { result } = await run([...policyArgs, "--interactive", "true"]);
    expect(result.isErr()).toBe(true);
    if (result.isErr()) {
      expect(result.error.name).toBe("ActionInteractiveUnsupported");
      expect(result.error.message).toContain("Action arguments");
      expect(result.error.message).toContain("configured token provider");
      expect(result.error.message).not.toContain("Sign in separately");
    }
    expect(core.executeAction).not.toHaveBeenCalled();
    expect(M365TokenProvider.getAccessToken).not.toHaveBeenCalled();
  });

  it.each([["--interactive", "true"], ["--interactive=true"], ["-i", "true"], ["-i=true"]])(
    "CLI-ACTION-04: rejects explicit interactive mode in CI (%j)",
    async (...flags) => {
      vi.stubEnv("CI_ENABLED", "true");
      const { result } = await run([...policyArgs, ...flags]);
      expect(result.isErr()).toBe(true);
      if (result.isErr()) expect(result.error.name).toBe("ActionInteractiveUnsupported");
      expect(core.executeAction).not.toHaveBeenCalled();
    }
  );

  it("CLI-ACTION-04: accepts explicit noninteractive mode without forwarding it as an arg", async () => {
    vi.stubEnv("CI_ENABLED", "true");
    const { result } = await run([...policyArgs, "--interactive", "false"]);
    expect(result.isOk()).toBe(true);
    expect(core.executeAction).toHaveBeenCalledExactlyOnceWith(
      ActionNames.PublishPolicy,
      {
        titleId: "U_title",
        policyPath: path.resolve("policy.json"),
        callerId: "11111111-2222-4333-8444-555555555555",
      },
      { projectPath: process.cwd() }
    );
  });

  it.each([
    ["--scope"],
    ["--scope", ""],
    ["--scope="],
    ["--scope", "Personal", "--scope"],
    ["--scope", "Personal", "--scope", ""],
    ["--scope", "Personal", "--scope", "--folder", "."],
  ])("CLI-ACTION-03: rejects explicit valueless optional flags (%j)", async (...flags) => {
    const { result } = await run([ActionNames.Acquire, "--app-package-path", "app.zip", ...flags]);
    expect(result.isErr()).toBe(true);
    expect(core.executeAction).not.toHaveBeenCalled();
    expect(M365TokenProvider.getAccessToken).not.toHaveBeenCalled();
  });

  it.each([
    ["--scope=--debug"],
    ["--scope", "Personal", "--scope=--debug"],
    ["--folder=--scope"],
    ["--scope", "Personal", "--folder=--scope"],
    ["--debug=--scope"],
    ["--scope", "Personal", "--debug=--scope"],
    ["--folder"],
    ["--folder", ""],
    ["--scope=Personal=ignored"],
    ["--app-package-path=app.zip=ignored"],
    ["--folder=.=ignored"],
  ])("CLI-ACTION-03: rejects ambiguous or valueless flag values (%j)", async (...flags) => {
    const { result } = await run([ActionNames.Acquire, "--app-package-path", "app.zip", ...flags]);
    expect(result.isErr()).toBe(true);
    expect(core.executeAction).not.toHaveBeenCalled();
    expect(M365TokenProvider.getAccessToken).not.toHaveBeenCalled();
  });

  it("CLI-ACTION-05: displays outputs and native fields without serializing callbacks", async () => {
    vi.mocked(core.executeAction).mockResolvedValue({
      result: ok(new Map([["titleId", "U_title"]])),
      summaries: ["done"],
      report,
      onCancellation: () =>
        new UserError({ source: "test", name: "Cancelled", message: "Cancelled" }),
    });
    const { result } = await run(policyArgs);
    expect(result.isOk()).toBe(true);
    expect(printedResults()).toContainEqual({
      outputs: { titleId: "U_title" },
      summaries: ["done"],
      report,
    });
  });

  it("CLI-ACTION-09: preserves opaque business metadata and valid JSON without mutating results", async () => {
    const titleId = "T_01234567-89ab-4cde-8123-456789abcdef";
    const appId = "01234567-89ab-4cde-8123-456789abcdef";
    const etag = `"${Buffer.from("synthetic-policy-version-12345").toString("base64")}"`;
    const metadataReport: PolicyPublicationReport = {
      ...report,
      operationId: appId,
      titleId,
      policySha256: "0123456789abcdef".repeat(4),
      workIq: { state: "confirmed", etag, revision: "synthetic-revision" },
    };
    const snapshot = structuredClone(metadataReport);
    vi.mocked(core.executeAction).mockResolvedValue({
      result: ok(
        new Map([
          ["titleId", titleId],
          ["appId", appId],
        ])
      ),
      summaries: ["done"],
      report: metadataReport,
    });
    const { result } = await run(policyArgs);
    expect(result.isOk()).toBe(true);
    expect(() => printedResults()).not.toThrow();
    expect(printedResults()).toContainEqual({
      outputs: { titleId, appId },
      summaries: ["done"],
      report: metadataReport,
    });
    expect(metadataReport).toEqual(snapshot);
  });

  it("CLI-ACTION-09: redacts credential values and nested containers while retaining keys", async () => {
    const redacted = "<REDACTED:secret>";
    const credentialsReport = Object.assign({}, report, {
      diagnostics: {
        authorization: "Bearer fixture-access",
        connectionString: "fixture-connection",
        credentials: { value: "fixture-credential" },
        entries: [{ password: "fixture-password", cookie: "fixture-cookie" }],
      },
    });
    vi.mocked(core.executeAction).mockResolvedValue({
      result: ok(
        new Map([
          ["accessToken", "fixture-token"],
          ["clientSecret", "fixture-secret"],
          ["apiKey", "fixture-key"],
        ])
      ),
      summaries: [],
      report: credentialsReport,
    });
    const { result } = await run(policyArgs);
    expect(result.isOk()).toBe(true);
    expect(() => printedResults()).not.toThrow();
    expect(printedResults()).toContainEqual({
      outputs: { accessToken: redacted, clientSecret: redacted, apiKey: redacted },
      summaries: [],
      report: {
        ...report,
        diagnostics: {
          authorization: redacted,
          connectionString: redacted,
          credentials: redacted,
          entries: [{ password: redacted, cookie: redacted }],
        },
      },
    });
  });

  it("CLI-ACTION-09: still masks registered environment secrets inside business metadata", async () => {
    const secret = "synthetic-sensitive-value";
    vi.stubEnv("SECRET_ACTION_OUTPUT_FIXTURE", secret);
    vi.mocked(core.executeAction).mockResolvedValue({
      result: ok(new Map([["titleId", secret]])),
      summaries: [],
      report: {
        ...report,
        titleId: secret,
        workIq: { state: "confirmed", etag: `"${secret}"` },
      },
    });
    const { result } = await run(policyArgs);
    expect(result.isOk()).toBe(true);
    expect(() => printedResults()).not.toThrow();
    expect(printedResults()).toContainEqual({
      outputs: { titleId: "<REDACTED:secret>" },
      summaries: [],
      report: {
        ...report,
        titleId: "<REDACTED:secret>",
        workIq: { state: "confirmed", etag: '"<REDACTED:secret>"' },
      },
    });
    expect(vi.mocked(logger.info).mock.calls.flat().join("\n")).not.toContain(secret);
  });

  it("CLI-ACTION-05: preserves failure receipts and the original error", async () => {
    const error = new UserError({ source: "test", name: "PublicationFailed", message: "Failed" });
    const failedReport: PolicyPublicationReport = {
      ...report,
      outcome: "incomplete",
      nextAction: "complete-publication",
      workIq: { state: "failed" },
    };
    vi.mocked(core.executeAction).mockResolvedValue({
      result: err(error),
      summaries: ["MOS confirmed"],
      report: failedReport,
    });
    const { result } = await run(policyArgs);
    expect(result.isErr()).toBe(true);
    if (result.isErr()) expect(result.error).toBe(error);
    expect(printedResults()).toContainEqual({
      summaries: ["MOS confirmed"],
      report: failedReport,
    });
  });

  it("SCN-CLI-ACTION-01: executes the real SDK and policy operation with fake destinations", async () => {
    vi.stubEnv("TEAMSFX_WORKIQ_PUBLISH_TARGET", "both");
    vi.mocked(core.executeAction).mockRestore();
    const mosPublish = vi.fn<() => Promise<PolicyDestinationResult>>();
    mosPublish.mockResolvedValue(ok({ state: "confirmed", revision: "mos-revision" }));
    const workIqPublish = vi.fn<() => Promise<PolicyDestinationResult>>();
    workIqPublish.mockResolvedValue(ok({ state: "confirmed", revision: "workiq-revision" }));
    const mos = vi
      .spyOn(PackageService.prototype, "prepareWorkIqPolicy")
      .mockResolvedValue(ok({ publish: mosPublish }));
    const workIq = vi
      .spyOn(workIqPolicyClient, "prepare")
      .mockResolvedValue(ok({ publish: workIqPublish }));
    const { result } = await run([
      ...policyArgs,
      "--policy-path",
      path.join(projectPath, "policy.json"),
    ]);
    if (result.isErr()) throw result.error;
    expect(result.isOk()).toBe(true);
    expect(mos).toHaveBeenCalledOnce();
    expect(workIq).toHaveBeenCalledOnce();
    expect(mos.mock.calls[0][0].policy.bytes.toString()).toBe('{"rules":[]}');
    expect(workIq.mock.calls[0][0].policy).toBe(mos.mock.calls[0][0].policy);
    expect(mosPublish).toHaveBeenCalledOnce();
    expect(workIqPublish).toHaveBeenCalledOnce();
    expect(printedResults()).toContainEqual(
      expect.objectContaining({
        outputs: {},
        report: expect.objectContaining({
          target: "both",
          outcome: "published",
          mos: { state: "confirmed", revision: "mos-revision" },
          workIq: { state: "confirmed", revision: "workiq-revision" },
        }),
      })
    );
    expect(await readdir(projectPath)).toEqual(["policy.json"]);
    expect(M365TokenProvider.getAccessToken).not.toHaveBeenCalled();
    expect(core.projectVersionCheck).not.toHaveBeenCalled();
  });
});
