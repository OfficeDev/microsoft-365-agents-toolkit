// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { CLIContext } from "@microsoft/teamsfx-api";
import { ActionNames } from "@microsoft/teamsfx-core/actions";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { engine } from "../../../src/commands/engine";
import { helper } from "../../../src/commands/helper";

describe("execute action feature flag", () => {
  beforeEach(() => {
    vi.resetModules();
    vi.stubEnv("TEAMSFX_CLI_BIN_NAME", "atk");
  });

  afterEach(() => {
    vi.unstubAllEnvs();
  });

  it.each([undefined, "false", "0"])(
    "CLI-ACTION-10: omits and rejects the whole execute tree when the flag is %s",
    async (value) => {
      vi.stubEnv("TEAMSFX_CLI_EXECUTE_ACTION", value);
      const { rootCommand } = await import("../../../src/commands/models/root");
      expect(rootCommand.commands?.some((command) => command.name === "execute")).toBe(false);
      expect(helper.formatHelp(rootCommand)).not.toMatch(/^\s+execute(?:\s|$)/m);
      expect(engine.findCommand(rootCommand, ["provision"]).cmd.name).toBe("provision");

      for (const args of [
        ["execute"],
        ["execute", "--help"],
        ["execute", "action", "--help"],
        ...Object.values(ActionNames).map((name) => ["execute", "action", name, "--help"]),
      ]) {
        const found = engine.findCommand(rootCommand, args);
        const context: CLIContext = {
          command: found.cmd,
          optionValues: {},
          globalOptionValues: {},
          argumentValues: [],
          telemetryProperties: {},
        };
        expect(found.cmd.fullName).toBe("atk");
        const result = engine.parseArgs(context, rootCommand, found.remainingArgs);
        expect(result.isErr()).toBe(true);
        if (result.isErr()) expect(result.error.message).toContain("'execute' is misspelled");
      }
    }
  );

  it.each(["true", "1", "TRUE"])(
    "CLI-ACTION-10: registers and lists the whole execute tree when the flag is %s",
    async (value) => {
      vi.stubEnv("TEAMSFX_CLI_EXECUTE_ACTION", value);
      const { rootCommand } = await import("../../../src/commands/models/root");
      expect(helper.formatHelp(rootCommand)).toMatch(/^\s+execute(?:\s|$)/m);
      expect(engine.findCommand(rootCommand, ["provision"]).cmd.name).toBe("provision");
      for (const name of Object.values(ActionNames)) {
        const found = engine.findCommand(rootCommand, ["execute", "action", name]);
        expect(found.cmd.fullName).toBe(`atk execute action ${name}`);
        expect(found.cmd.handler).toBeTypeOf("function");
        expect(found.remainingArgs).toEqual([]);
      }
    }
  );
});
