// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { SystemError } from "@microsoft/teamsfx-api";
import {
  STEP_MATERIALIZE_STATIC_MCP_TOOLS,
  mcpStaticMaterializeTools,
} from "../../../../src/v4/runtime/steps/mcpStatic";
import { StepContext } from "../../../../src/v4/pipeline/runScaffoldPipeline";
import { assert } from "vitest";

function makeCtx(initial: Record<string, string> = {}): {
  ctx: StepContext;
  files: Map<string, Buffer>;
} {
  const files = new Map<string, Buffer>();
  for (const [filePath, body] of Object.entries(initial)) {
    files.set(filePath, Buffer.from(body, "utf8"));
  }
  return {
    files,
    ctx: {
      read: (filePath) => files.get(filePath),
      write: (filePath, data) => {
        files.set(filePath, data);
      },
      manifestWrapper: () => ({ addAction: () => undefined }),
    },
  };
}

function validParams(overrides: Record<string, unknown> = {}): Record<string, unknown> {
  return {
    pluginPath: "appPackage/ai-plugin.json",
    toolsPath: "appPackage/mcp-tools-1.json",
    mcpServerUrl: "https://api.example.com/mcp",
    toolsJson: JSON.stringify({
      tools: [
        { name: "searchFlights", description: "Search flights" },
        { name: "bookFlight", description: "Book flights" },
      ],
    }),
    selected: ["searchFlights"],
    ...overrides,
  };
}

describe(`${STEP_MATERIALIZE_STATIC_MCP_TOOLS} (v4)`, () => {
  it("validateParams reports missing or invalid parameters", () => {
    assert.strictEqual(
      mcpStaticMaterializeTools.validateParams(validParams({ pluginPath: undefined })),
      "missing string parameter 'pluginPath'"
    );
    assert.strictEqual(
      mcpStaticMaterializeTools.validateParams(validParams({ toolsPath: undefined })),
      "missing string parameter 'toolsPath'"
    );
    assert.strictEqual(
      mcpStaticMaterializeTools.validateParams(validParams({ mcpServerUrl: undefined })),
      "missing string parameter 'mcpServerUrl'"
    );
    assert.strictEqual(
      mcpStaticMaterializeTools.validateParams(validParams({ toolsJson: undefined })),
      "missing string parameter 'toolsJson'"
    );
    assert.strictEqual(
      mcpStaticMaterializeTools.validateParams(validParams({ selected: ["ok", 1] })),
      "missing string[] parameter 'selected'"
    );
  });

  it("returns a parameter SystemError when apply receives invalid resolved params", () => {
    const { ctx } = makeCtx();
    const result = mcpStaticMaterializeTools.apply(validParams({ selected: "searchFlights" }), ctx);

    assert.isTrue(result.isErr());
    assert.instanceOf(result._unsafeUnwrapErr(), SystemError);
    assert.strictEqual(result._unsafeUnwrapErr().name, "McpStaticParams");
  });

  it("returns tool parser and selection errors as SystemError results", () => {
    const { ctx } = makeCtx({ "appPackage/ai-plugin.json": "{}" });

    const parseResult = mcpStaticMaterializeTools.apply(
      validParams({ toolsJson: "not json" }),
      ctx
    );
    assert.isTrue(parseResult.isErr());
    assert.strictEqual(parseResult._unsafeUnwrapErr().name, "McpStaticToolsParse");

    const missingSelection = mcpStaticMaterializeTools.apply(
      validParams({ selected: ["missingTool"] }),
      ctx
    );
    assert.isTrue(missingSelection.isErr());
    assert.strictEqual(missingSelection._unsafeUnwrapErr().name, "McpStaticToolMissing");
  });

  it("returns plugin read and parse errors as SystemError results", () => {
    const missing = mcpStaticMaterializeTools.apply(validParams(), makeCtx().ctx);
    assert.isTrue(missing.isErr());
    assert.strictEqual(missing._unsafeUnwrapErr().name, "McpStaticPluginMissing");

    const invalidJson = mcpStaticMaterializeTools.apply(
      validParams(),
      makeCtx({ "appPackage/ai-plugin.json": "not json" }).ctx
    );
    assert.isTrue(invalidJson.isErr());
    assert.strictEqual(invalidJson._unsafeUnwrapErr().name, "McpStaticPluginParse");

    const invalidShape = mcpStaticMaterializeTools.apply(
      validParams(),
      makeCtx({ "appPackage/ai-plugin.json": "[]" }).ctx
    );
    assert.isTrue(invalidShape.isErr());
    assert.strictEqual(invalidShape._unsafeUnwrapErr().name, "McpStaticPluginShape");
  });

  it("writes selected tools and RemoteMCPServer runtime metadata", () => {
    const { ctx, files } = makeCtx({ "appPackage/ai-plugin.json": "{}" });

    const result = mcpStaticMaterializeTools.apply(
      validParams({ toolsPath: "nested/appPackage/mcp-tools-1.json" }),
      ctx
    );

    assert.isTrue(result.isOk(), result.isErr() ? result.error.message : "expected ok");
    const plugin = JSON.parse(files.get("appPackage/ai-plugin.json")?.toString("utf8") ?? "{}");
    assert.deepEqual(plugin.functions, [{ name: "searchFlights", description: "Search flights" }]);
    assert.deepEqual(plugin.runtimes, [
      {
        type: "RemoteMCPServer",
        spec: {
          url: "https://api.example.com/mcp",
          mcp_tool_description: { file: "mcp-tools-1.json" },
        },
        run_for_functions: ["searchFlights"],
      },
    ]);
    const tools = JSON.parse(
      files.get("nested/appPackage/mcp-tools-1.json")?.toString("utf8") ?? "{}"
    );
    assert.deepEqual(tools.tools, [{ name: "searchFlights", description: "Search flights" }]);
  });
});
