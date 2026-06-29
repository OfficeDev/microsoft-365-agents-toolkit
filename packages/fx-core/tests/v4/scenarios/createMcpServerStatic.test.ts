// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { assert } from "chai";
import { UserError } from "@microsoft/teamsfx-api";
import { REQUIRE_EMPTY_TARGET } from "../../../src/v4/pipeline/runScaffoldPipeline";
import { createInMemoryRuntime } from "../../../src/v4/runtime/inMemoryRuntime";
import { ScaffoldRequest, scaffold } from "../../../src/v4/runtime/scaffold";
import {
  loadV4Package,
  readJsonObject,
  recordArrayProperty,
  recordProperty,
  runV4Package,
  V4ScenarioOutcome,
} from "./helpers/scenarioHarness";

/**
 * T3 scenario tier for the DT-off v4 static MCP create package.
 *
 * Spec: docs/03-specs/scenarios/da/create-mcp-server-static.md
 * (SCN-CREATE-MCP-STATIC-01..04)
 */

const MCP_SERVER_URL = "https://api.github.com/mcp";
const MCP_TOOLS_JSON = JSON.stringify({
  tools: [
    {
      name: "search",
      description: "Search issues",
      inputSchema: { type: "object", properties: { query: { type: "string" } } },
    },
    {
      name: "calendar",
      description: "Read calendar",
      inputSchema: { type: "object", properties: { date: { type: "string" } } },
    },
  ],
});

const templatePackage = loadV4Package("create", "da/mcp-server-static");

async function run(
  selectedMcpTools: string[] = ["search"]
): Promise<{ files: Map<string, Buffer>; outcome: V4ScenarioOutcome }> {
  return runV4Package(templatePackage, {
    answers: {
      mcpServerUrl: MCP_SERVER_URL,
      mcpToolsJson: MCP_TOOLS_JSON,
      selectedMcpTools,
    },
    callerFloor: { appName: "MyStaticMcpAgent", language: "common" },
  });
}

describe("SCN-DA-CREATE-WITH-MCP-SERVER-STATIC (v4, T3 InMemoryRuntime)", () => {
  it("SCN-CREATE-MCP-STATIC-01: static MCP scaffold writes mcp-tools-1.json", async () => {
    const { files, outcome } = await run(["search", "calendar"]);
    assert.include(outcome.written, "appPackage/ai-plugin.json");
    assert.include(outcome.written, "appPackage/declarativeAgent.json");
    assert.include(outcome.written, "appPackage/manifest.json");
    assert.include(outcome.stepsRun, "mcp-static/materialize-tools");
    assert.isTrue(files.has("appPackage/mcp-tools-1.json"));
  });

  it("SCN-CREATE-MCP-STATIC-02: selected tools filter functions and full tool definitions", async () => {
    const { files } = await run(["search"]);
    const plugin = readJsonObject(files, "appPackage/ai-plugin.json");
    const functions = recordArrayProperty(plugin, "functions");
    assert.deepStrictEqual(functions, [{ name: "search", description: "Search issues" }]);

    const toolsFile = readJsonObject(files, "appPackage/mcp-tools-1.json");
    const tools = recordArrayProperty(toolsFile, "tools");
    assert.lengthOf(tools, 1);
    assert.strictEqual(tools[0].name, "search");
    assert.deepStrictEqual(recordProperty(tools[0], "inputSchema"), {
      type: "object",
      properties: { query: { type: "string" } },
    });
  });

  it("SCN-CREATE-MCP-STATIC-03: static runtime references mcp-tools and omits dynamic discovery", async () => {
    const { files } = await run(["search"]);
    const plugin = readJsonObject(files, "appPackage/ai-plugin.json");
    const runtime = recordArrayProperty(plugin, "runtimes")[0];
    const spec = recordProperty(runtime, "spec");
    const description = recordProperty(spec, "mcp_tool_description");
    assert.strictEqual(runtime.type, "RemoteMCPServer");
    assert.strictEqual(spec.url, MCP_SERVER_URL);
    assert.strictEqual(description.file, "mcp-tools-1.json");
    assert.deepStrictEqual(runtime.run_for_functions, ["search"]);
    assert.notProperty(spec, "enable_dynamic_discovery");
  });

  it("SCN-CREATE-MCP-STATIC-04: non-empty target fails require-empty-target first", async () => {
    const runtime = createInMemoryRuntime();
    const request: ScaffoldRequest = {
      descriptor: templatePackage.descriptor,
      pipeline: templatePackage.pipeline,
      content: templatePackage.content,
      answers: {
        mcpServerUrl: MCP_SERVER_URL,
        mcpToolsJson: MCP_TOOLS_JSON,
        selectedMcpTools: ["search"],
      },
      callerFloor: { appName: "MyStaticMcpAgent", language: "common" },
      targetDir: { path: "/out", existing: ["appPackage/manifest.json"] },
    };
    const result = await scaffold(request, runtime);
    assert.isTrue(result.isErr());
    const error = result._unsafeUnwrapErr();
    assert.instanceOf(error, UserError);
    assert.strictEqual(error.name, REQUIRE_EMPTY_TARGET);
    assert.strictEqual(runtime.files.size, 0);
  });
});
