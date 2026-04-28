// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Zhiyu You <zhiyou@microsoft.com>
 */

import { ProgrammingLanguage } from "@microsoft/teamsfx-core";
import { expect } from "chai";
import * as fs from "fs-extra";
import * as path from "path";
import { Capability } from "../../../utils/constants";
import { CaseFactory } from "../../caseFactory";
import {
  mcpToolsFilePath,
  writeMCPToolsFixture,
  removeMCPToolsFixture,
} from "./mcpToolsFixture";

// Case 1: atk new — MCP with no-auth server URL (auto-fetch tools)
class DeclarativeAgentMCPNoAuthNew extends CaseFactory {
  public override async onAfter(projectPath: string): Promise<void> {
    await fs.remove(projectPath);
  }

  public override async onAfterCreate(projectPath: string): Promise<void> {
    const appPackage = path.join(projectPath, "appPackage");

    // ai-plugin.json must exist and have functions and MCP runtime
    const aiPlugin = await fs.readJSON(path.join(appPackage, "ai-plugin.json"));
    expect(aiPlugin.functions).to.be.an("array").that.is.not.empty;
    expect(aiPlugin.runtimes).to.be.an("array").that.is.not.empty;
    expect(aiPlugin.runtimes[0].type).to.equal("RemoteMCPServer");
    expect(aiPlugin.runtimes[0].spec.url).to.be.a("string").that.is.not.empty;
    // No auth block for no-auth servers
    expect(aiPlugin.runtimes[0].auth).to.be.undefined;

    // mcp-tools-1.json must exist with tool definitions
    const mcpToolsPath = path.join(appPackage, "mcp-tools-1.json");
    expect(fs.pathExistsSync(mcpToolsPath)).to.be.true;
    const mcpTools = await fs.readJSON(mcpToolsPath);
    expect(mcpTools.tools).to.be.an("array").that.is.not.empty;

    // DA manifest must reference ai-plugin.json
    const daManifest = await fs.readJSON(
      path.join(appPackage, "declarativeAgent.json"),
    );
    expect(daManifest.actions).to.be.an("array").that.is.not.empty;
    expect(daManifest.actions[0].file).to.equal("ai-plugin.json");

    // m365agents.yml should NOT contain oauth/register action
    const ymlPath = path.join(projectPath, "m365agents.yml");
    if (fs.pathExistsSync(ymlPath)) {
      const ymlContent = fs.readFileSync(ymlPath, "utf8");
      expect(ymlContent).to.not.include("oauth/register");
    }
  }
}

// Case 6: atk new — MCP with tools loaded from file (no auth)
class DeclarativeAgentMCPNoAuthFile extends CaseFactory {
  public override async onBefore(): Promise<void> {
    await writeMCPToolsFixture();
  }

  public override async onAfter(projectPath: string): Promise<void> {
    await fs.remove(projectPath);
    await removeMCPToolsFixture();
  }

  public override async onAfterCreate(projectPath: string): Promise<void> {
    const appPackage = path.join(projectPath, "appPackage");

    // ai-plugin.json must exist and have functions and MCP runtime
    const aiPlugin = await fs.readJSON(path.join(appPackage, "ai-plugin.json"));
    expect(aiPlugin.functions).to.be.an("array").that.is.not.empty;
    expect(aiPlugin.runtimes).to.be.an("array").that.is.not.empty;
    expect(aiPlugin.runtimes[0].type).to.equal("RemoteMCPServer");
    // No auth block for no-auth servers
    expect(aiPlugin.runtimes[0].auth).to.be.undefined;

    // mcp-tools-1.json must exist
    const mcpToolsPath = path.join(appPackage, "mcp-tools-1.json");
    expect(fs.pathExistsSync(mcpToolsPath)).to.be.true;
    const mcpTools = await fs.readJSON(mcpToolsPath);
    expect(mcpTools.tools).to.be.an("array").that.is.not.empty;
  }
}

// Case 1: No-auth server URL — auto-fetch
const noAuthUrlRecord: Record<string, string> = {};
noAuthUrlRecord["with-plugin"] = "yes";
noAuthUrlRecord["api-plugin-type"] = "mcp";
noAuthUrlRecord["mcp-da-server-url"] = "https://learn.microsoft.com/api/mcp";

new DeclarativeAgentMCPNoAuthNew(
  Capability.DeclarativeAgent,
  37357419,
  "zhiyou@microsoft.com",
  [],
  ProgrammingLanguage.None,
  {
    skipProvision: true,
  },
  noAuthUrlRecord,
).test();

// Case 6: No-auth with tools from file
const noAuthFileRecord: Record<string, string> = {};
noAuthFileRecord["with-plugin"] = "yes";
noAuthFileRecord["api-plugin-type"] = "mcp";
noAuthFileRecord["mcp-da-server-url"] = "https://learn.microsoft.com/api/mcp";
noAuthFileRecord["mcp-tools-file-path"] = mcpToolsFilePath;

new DeclarativeAgentMCPNoAuthFile(
  Capability.DeclarativeAgent,
  37357445,
  "zhiyou@microsoft.com",
  [],
  ProgrammingLanguage.None,
  {
    skipProvision: true,
  },
  noAuthFileRecord,
).test();
