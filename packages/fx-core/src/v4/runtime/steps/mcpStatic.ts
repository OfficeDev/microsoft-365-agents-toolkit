// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { FxError, SystemError } from "@microsoft/teamsfx-api";
import { Result, err, ok } from "neverthrow";
import { parseMcpStaticToolsJson, selectMcpStaticTools } from "../../mcp/mcpStaticTools";
import { RegisteredStep, StepContext, StepParams } from "../../pipeline/runScaffoldPipeline";

/** Materialize static MCP tools and plugin runtime metadata. */

const SOURCE = "Scaffold";

/** Engine step name `mcp-static/materialize-tools`. */
export const STEP_MATERIALIZE_STATIC_MCP_TOOLS = "mcp-static/materialize-tools";

function systemError(name: string, message: string): SystemError {
  return new SystemError({ source: SOURCE, name, message });
}

function stringParam(params: StepParams, key: string): string | undefined {
  const value = params[key];
  return typeof value === "string" ? value : undefined;
}

function stringArrayParam(params: StepParams, key: string): string[] | undefined {
  const value = params[key];
  if (!Array.isArray(value) || !value.every((item) => typeof item === "string")) {
    return undefined;
  }
  return value;
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function errorMessage(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}

function readRequiredJson(
  ctx: StepContext,
  pluginPath: string
): Result<Record<string, unknown>, FxError> {
  const current = ctx.read(pluginPath);
  if (current === undefined) {
    return err(
      systemError(
        "McpStaticPluginMissing",
        `Cannot materialize static MCP tools because '${pluginPath}' was not produced by the render phase.`
      )
    );
  }
  let parsed: unknown;
  try {
    parsed = JSON.parse(current.toString("utf8"));
  } catch (error) {
    return err(
      systemError(
        "McpStaticPluginParse",
        `Cannot parse '${pluginPath}' as JSON: ${errorMessage(error)}`
      )
    );
  }
  if (!isRecord(parsed)) {
    return err(systemError("McpStaticPluginShape", `'${pluginPath}' is not a JSON object`));
  }
  return ok(parsed);
}

function fileName(filePath: string): string {
  const segments = filePath.split("/").filter((segment) => segment.length > 0);
  return segments[segments.length - 1] ?? filePath;
}

/** Registered step for writing `mcp-tools-1.json` and static RemoteMCPServer metadata. */
export const mcpStaticMaterializeTools: RegisteredStep = {
  validateParams(resolved: StepParams): string | undefined {
    if (stringParam(resolved, "pluginPath") === undefined) {
      return "missing string parameter 'pluginPath'";
    }
    if (stringParam(resolved, "toolsPath") === undefined) {
      return "missing string parameter 'toolsPath'";
    }
    if (stringParam(resolved, "mcpServerUrl") === undefined) {
      return "missing string parameter 'mcpServerUrl'";
    }
    if (stringParam(resolved, "toolsJson") === undefined) {
      return "missing string parameter 'toolsJson'";
    }
    if (stringArrayParam(resolved, "selected") === undefined) {
      return "missing string[] parameter 'selected'";
    }
    return undefined;
  },
  apply(resolved: StepParams, ctx: StepContext): Result<void, FxError> {
    const pluginPath = stringParam(resolved, "pluginPath");
    const toolsPath = stringParam(resolved, "toolsPath");
    const mcpServerUrl = stringParam(resolved, "mcpServerUrl");
    const toolsJson = stringParam(resolved, "toolsJson");
    const selected = stringArrayParam(resolved, "selected");
    if (
      pluginPath === undefined ||
      toolsPath === undefined ||
      mcpServerUrl === undefined ||
      toolsJson === undefined ||
      selected === undefined
    ) {
      return err(
        systemError("McpStaticParams", "resolved parameters are not all of the expected type")
      );
    }

    const parsedTools = parseMcpStaticToolsJson(toolsJson);
    if (!parsedTools.ok) {
      return err(systemError(parsedTools.code, parsedTools.message));
    }
    const selectedTools = selectMcpStaticTools(parsedTools.tools, selected);
    if (!selectedTools.ok) {
      return err(systemError(selectedTools.code, selectedTools.message));
    }

    const plugin = readRequiredJson(ctx, pluginPath);
    if (plugin.isErr()) {
      return err(plugin.error);
    }

    plugin.value.functions = selectedTools.tools.map((tool) => ({
      name: tool.name,
      description: tool.description,
    }));
    plugin.value.runtimes = [
      {
        type: "RemoteMCPServer",
        spec: {
          url: mcpServerUrl,
          mcp_tool_description: {
            file: fileName(toolsPath),
          },
        },
        run_for_functions: selectedTools.tools.map((tool) => tool.name),
      },
    ];

    ctx.write(pluginPath, Buffer.from(JSON.stringify(plugin.value, null, 4) + "\n", "utf8"));
    ctx.write(
      toolsPath,
      Buffer.from(
        JSON.stringify({ tools: selectedTools.tools.map((tool) => tool.raw) }, null, 4) + "\n",
        "utf8"
      )
    );
    return ok(undefined);
  },
};
