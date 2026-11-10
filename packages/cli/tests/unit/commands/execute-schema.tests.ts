// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { CLIContext, UserError, err, ok } from "@microsoft/teamsfx-api";
import * as actions from "@microsoft/teamsfx-core/actions";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { getFxCore } from "../../../src/activate";
import { getExecuteCommand } from "../../../src/commands/models/execute";
import { rootCommand } from "../../../src/commands/models/root";

vi.mock("@microsoft/teamsfx-core/actions", async (importOriginal) => {
  const actual = await importOriginal<typeof actions>();
  return { ...actual, getActionDefinition: vi.fn(actual.getActionDefinition) };
});

describe("action schema adapter", () => {
  beforeEach(async () => {
    const actual = await vi.importActual<typeof actions>("@microsoft/teamsfx-core/actions");
    vi.mocked(actions.getActionDefinition).mockImplementation(actual.getActionDefinition);
    vi.spyOn(getFxCore(), "executeAction");
  });

  afterEach(() => {
    vi.restoreAllMocks();
  });

  async function runPolicyCommand() {
    const command = getExecuteCommand().commands?.[0].commands?.find(
      (candidate) => candidate.name === actions.ActionNames.PublishPolicy
    );
    expect(command?.handler).toBeTypeOf("function");
    if (!command?.handler) throw new Error("Missing action handler");
    const context: CLIContext = {
      command: { ...command, fullName: `atk execute action ${command.name}` },
      argumentValues: [],
      optionValues: {},
      globalOptionValues: { interactive: false },
      telemetryProperties: {},
    };
    return await command.handler(context);
  }

  it.each([
    { type: "array" },
    { type: "object", properties: { item: { type: "object" } } },
    { type: "object", properties: { item: { type: "string", enum: ["a", 1] } } },
    { type: "object", properties: { folder: { type: "string" } } },
    {
      type: "object",
      properties: { fooBar: { type: "string" }, "foo-bar": { type: "string" } },
    },
    { type: "object", properties: { item: { type: "string" } }, required: "item" },
  ])("CLI-ACTION-06: reports unsupported or ambiguous schema (%j)", async (argsSchema) => {
    vi.mocked(actions.getActionDefinition).mockReturnValue(
      ok({
        name: actions.ActionNames.PublishPolicy,
        argsSchema,
        supportsCancellation: true,
      })
    );
    const result = await runPolicyCommand();
    expect(result.isErr()).toBe(true);
    if (result.isErr()) expect(result.error.name).toBe("UnsupportedActionSchema");
    expect(getFxCore().executeAction).not.toHaveBeenCalled();
    expect(rootCommand.commands?.some((command) => command.name === "install")).toBe(true);
  });

  it("CLI-ACTION-06: preserves metadata lookup errors", async () => {
    const error = new UserError({ source: "test", name: "MetadataFailed", message: "Failed" });
    vi.mocked(actions.getActionDefinition).mockReturnValue(err(error));
    const result = await runPolicyCommand();
    expect(result.isErr()).toBe(true);
    if (result.isErr()) expect(result.error).toBe(error);
    expect(getFxCore().executeAction).not.toHaveBeenCalled();
  });
});
