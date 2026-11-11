// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import AdmZip from "adm-zip";
import { M365TokenProvider, ok } from "@microsoft/teamsfx-api";
import { vi } from "vitest";
export const productId = "11111111-2222-4333-8444-555555555555";
export const appId = "aaaaaaaa-bbbb-4ccc-8ddd-eeeeeeeeeeee";
export const componentId = "bbbbbbbb-cccc-4ddd-8eee-ffffffffffff";
export const tenantUrl = "https://tenant.example.test";
export const tenantId = "cccccccc-dddd-4eee-8fff-000000000000";
export const publicationTenantContext = {
  tenantId: "cccccccc-dddd-4eee-8fff-000000000000",
  tenantUrl,
  catalogUrl: `${tenantUrl}/sites/appcatalog`,
  productId,
};

export function createSppkg(version = "1.2.3.4", componentContent = "component"): Buffer {
  const agent = new AdmZip();
  agent.addFile(
    "ai-plugin.json",
    Buffer.from(
      JSON.stringify({
        schema_version: "v2.4",
        namespace: "sample",
        name_for_human: "Sample",
        description_for_human: "Sample API plugin",
        functions: [],
        runtimes: [
          {
            type: "RemoteMCPServer",
            auth: { type: "None" },
            spec: { url: "{{TENANT_MCP_URL}}" },
          },
        ],
      })
    )
  );
  agent.addFile("manifest.json", Buffer.from("{}"));
  agent.addFile("declarativeAgent.json", Buffer.from("{}"));

  const packageZip = new AdmZip();
  packageZip.addFile(
    "AppManifest.xml",
    Buffer.from(
      `<App ProductID="{${productId}}" Version="${version}" IsClientSideSolution="true" SkipFeatureDeployment="false" />`
    )
  );
  packageZip.addFile(
    "[Content_Types].xml",
    Buffer.from(
      `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="text/xml" /><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml" /><Default Extension="zip" ContentType="application/octet-stream" /></Types>`
    )
  );
  packageZip.addFile(
    "_rels/.rels",
    Buffer.from(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Type="manifest" Target="/AppManifest.xml" Id="r0" /></Relationships>`
    )
  );
  packageZip.addFile(
    "_rels/ClientSideAssets.xml.rels",
    Buffer.from(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Type="agent" Target="/ClientSideAssets/agent.zip" Id="agent" /><Relationship Type="component" Target="/ClientSideAssets/component.js" Id="component" /></Relationships>`
    )
  );
  packageZip.addFile("ClientSideAssets.xml", Buffer.from("<Assets />"));
  packageZip.addFile(
    `${componentId}/CopilotComponent_${componentId}.xml`,
    Buffer.from("<Elements><ClientSideComponent /></Elements>")
  );
  packageZip.addFile("ClientSideAssets/agent.zip", agent.toBuffer());
  packageZip.addFile("ClientSideAssets/component.js", Buffer.from(componentContent));
  return packageZip.toBuffer();
}

export function createAppPackage(version = "2.3.4"): Buffer {
  const zip = new AdmZip();
  zip.addFile(
    "manifest.json",
    Buffer.from(
      JSON.stringify({
        manifestVersion: "1.30",
        version,
        id: appId,
        $schema:
          "https://developer.microsoft.com/json-schemas/teams/v1.30/MicrosoftTeams.schema.json",
        developer: {
          name: "Sample developer",
          websiteUrl: "https://example.test",
          privacyUrl: "https://example.test/privacy",
          termsOfUseUrl: "https://example.test/terms",
        },
        name: { short: "Sample app", full: "Sample application" },
        description: { short: "Sample app", full: "Sample application" },
        icons: { color: "color.png", outline: "outline.png" },
        accentColor: "#FFFFFF",
        copilotAgents: {
          declarativeAgents: [{ id: "agent", file: "declarativeAgent.json" }],
        },
      })
    )
  );
  zip.addFile(
    "declarativeAgent.json",
    Buffer.from(
      JSON.stringify({
        version: "v1.9",
        name: "Sample agent",
        description: "Sample agent description",
        instructions: "Keep these instructions",
        actions: [],
      })
    )
  );
  zip.addFile("color.png", Buffer.from("icon"));
  zip.addFile("outline.png", Buffer.from("icon"));
  return zip.toBuffer();
}

export interface TokenClaimsOverride {
  graphTenantId?: string;
  graphAudience?: string;
  sharePointTenantId?: string;
  sharePointAudience?: string;
}

/** Returns signed-in Graph and SharePoint claims for the fixture tenant unless overridden. */
export function mockTokenClaims(
  provider: M365TokenProvider,
  override: TokenClaimsOverride = {}
): void {
  vi.spyOn(provider, "getJsonObject").mockImplementation(async (tokenRequest) =>
    ok(
      tokenRequest.scopes.some((scope) => scope.includes("Sites.FullControl.All"))
        ? {
            tid: override.sharePointTenantId ?? override.graphTenantId ?? tenantId,
            aud: override.sharePointAudience ?? tenantUrl,
          }
        : {
            tid: override.graphTenantId ?? tenantId,
            aud: override.graphAudience ?? "https://graph.microsoft.com",
          }
    )
  );
}
