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

// Verification for MCP projects when --mcp-da-auth-type is specified.
// Note: learn.microsoft.com/api/mcp is a public no-auth server, so even when
// --mcp-da-auth-type is passed, the server probe detects no auth requirement
// and the project is scaffolded without auth blocks. This is correct behavior —
// auth is driven by server probe, not solely by the CLI flag.
// To test actual auth injection, an auth-required MCP server is needed.
class DeclarativeAgentMCPWithAuth extends CaseFactory {
  private authType: "oauth" | "entraSSO";

  public override async onBefore(): Promise<void> {
    await writeMCPToolsFixture();
  }

  public override async onAfter(projectPath: string): Promise<void> {
    await fs.remove(projectPath);
    await removeMCPToolsFixture();
  }

  public constructor(
    authType: "oauth" | "entraSSO",
    testPlanCaseId: number,
    author: string,
    custimized: Record<string, string>,
  ) {
    super(
      Capability.DeclarativeAgent,
      testPlanCaseId,
      author,
      [],
      ProgrammingLanguage.None,
      { skipProvision: true },
      custimized,
    );
    this.authType = authType;
  }

  public override async onAfterCreate(projectPath: string): Promise<void> {
    const appPackage = path.join(projectPath, "appPackage");

    // ai-plugin.json must exist with MCP runtime
    const aiPlugin = await fs.readJSON(path.join(appPackage, "ai-plugin.json"));
    expect(aiPlugin.functions).to.be.an("array").that.is.not.empty;
    expect(aiPlugin.runtimes).to.be.an("array").that.is.not.empty;
    expect(aiPlugin.runtimes[0].type).to.equal("RemoteMCPServer");
    expect(aiPlugin.runtimes[0].spec.url).to.be.a("string").that.is.not.empty;

    // With a no-auth server, auth block should be absent even if --mcp-da-auth-type was specified
    // Auth injection is gated on server probe detecting auth requirement
    expect(aiPlugin.runtimes[0].auth).to.be.undefined;

    // mcp-tools-1.json must exist with tool definitions
    const mcpToolsPath = path.join(appPackage, "mcp-tools-1.json");
    expect(fs.pathExistsSync(mcpToolsPath)).to.be.true;
    const mcpTools = await fs.readJSON(mcpToolsPath);
    expect(mcpTools.tools).to.be.an("array").that.is.not.empty;

    // m365agents.yml should NOT contain oauth/register for a no-auth server
    const ymlPath = path.join(projectPath, "m365agents.yml");
    if (fs.pathExistsSync(ymlPath)) {
      const ymlContent = fs.readFileSync(ymlPath, "utf8");
      expect(ymlContent).to.not.include("oauth/register");
    }

    // DA manifest must reference ai-plugin.json
    const daManifest = await fs.readJSON(
      path.join(appPackage, "declarativeAgent.json"),
    );
    expect(daManifest.actions).to.be.an("array").that.is.not.empty;
    expect(daManifest.actions[0].file).to.equal("ai-plugin.json");
  }
}

// Case 3: atk new — MCP with OAuth auth (server URL, auth detected)
const oauthRecord: Record<string, string> = {};
oauthRecord["with-plugin"] = "yes";
oauthRecord["api-plugin-type"] = "mcp";
oauthRecord["mcp-da-server-url"] = "https://learn.microsoft.com/api/mcp";
oauthRecord["mcp-da-auth-type"] = "oauth";

new DeclarativeAgentMCPWithAuth(
  "oauth",
  37357426,
  "zhiyou@microsoft.com",
  oauthRecord,
).test();

// Case 4: atk new — MCP with EntraSSO auth
const entraRecord: Record<string, string> = {};
entraRecord["with-plugin"] = "yes";
entraRecord["api-plugin-type"] = "mcp";
entraRecord["mcp-da-server-url"] = "https://learn.microsoft.com/api/mcp";
entraRecord["mcp-da-auth-type"] = "entraSSO";

new DeclarativeAgentMCPWithAuth(
  "entraSSO",
  37357431,
  "zhiyou@microsoft.com",
  entraRecord,
).test();

// Case 8: atk new — MCP with OAuth auth + tools from file
const oauthFileRecord: Record<string, string> = {};
oauthFileRecord["with-plugin"] = "yes";
oauthFileRecord["api-plugin-type"] = "mcp";
oauthFileRecord["mcp-da-server-url"] = "https://learn.microsoft.com/api/mcp";
oauthFileRecord["mcp-da-auth-type"] = "oauth";
oauthFileRecord["mcp-tools-file-path"] = mcpToolsFilePath;

new DeclarativeAgentMCPWithAuth(
  "oauth",
  37357426,
  "zhiyou@microsoft.com",
  oauthFileRecord,
).test();

// Case 9: atk new — MCP with EntraSSO auth + tools from file
const entraFileRecord: Record<string, string> = {};
entraFileRecord["with-plugin"] = "yes";
entraFileRecord["api-plugin-type"] = "mcp";
entraFileRecord["mcp-da-server-url"] = "https://learn.microsoft.com/api/mcp";
entraFileRecord["mcp-da-auth-type"] = "entraSSO";
entraFileRecord["mcp-tools-file-path"] = mcpToolsFilePath;

new DeclarativeAgentMCPWithAuth(
  "entraSSO",
  37357431,
  "zhiyou@microsoft.com",
  entraFileRecord,
).test();
