// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import AdmZip from "adm-zip";
import FormData from "form-data";
import { err, FxError, M365TokenProvider, ok, Result, SystemError } from "@microsoft/teamsfx-api";
import { vi } from "vitest";
import { MosServiceScope } from "../../../../../src/common/constants";

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

/** Zip-level differences that a rebuild of an unchanged application package can introduce. */
export interface AppPackageRebuild {
  /** Modification time stamped on every entry, as a rebuild at another time would. */
  entryTime?: Date;
  reverseOrder?: boolean;
}

export function createAppPackage(version = "2.3.4", rebuild: AppPackageRebuild = {}): Buffer {
  const entries: [string, Buffer][] = [
    [
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
      ),
    ],
    [
      "declarativeAgent.json",
      Buffer.from(
        JSON.stringify({
          version: "v1.9",
          name: "Sample agent",
          description: "Sample agent description",
          instructions: "Keep these instructions",
          actions: [],
        })
      ),
    ],
    ["color.png", Buffer.from("icon")],
    ["outline.png", Buffer.from("icon")],
  ];
  if (rebuild.reverseOrder) entries.reverse();
  const zip = new AdmZip();
  for (const [name, content] of entries) {
    zip.addFile(name, content);
    const entry = zip.getEntry(name);
    if (entry && rebuild.entryTime) entry.header.time = rebuild.entryTime;
  }
  return zip.toBuffer();
}

export interface TokenClaimsOverride {
  graphTenantId?: string;
  graphAudience?: string;
  sharePointTenantId?: string;
  sharePointAudience?: string;
  mosTenantId?: string;
  mosAudience?: string;
  /** Makes the SharePoint claims unreadable, as when the signed-in token cannot be decoded. */
  sharePointClaimsUnavailable?: boolean;
  /** Makes the MOS3 claims unreadable, as when the signed-in token cannot be decoded. */
  mosClaimsUnavailable?: boolean;
}

function claimsUnavailable(): Result<Record<string, unknown>, FxError> {
  return err(
    new SystemError({
      source: "test",
      name: "ClaimsUnavailable",
      message: "The signed-in token claims are unavailable.",
    })
  );
}

/**
 * Returns signed-in Graph, SharePoint and MOS3 claims for the fixture tenant unless overridden.
 * The MOS3 claims are independent of the Graph claims because a prepared submission never
 * requests a Graph token.
 */
export function mockTokenClaims(
  provider: M365TokenProvider,
  override: TokenClaimsOverride = {}
): void {
  vi.spyOn(provider, "getJsonObject").mockImplementation(async (tokenRequest) => {
    if (tokenRequest.scopes.some((scope) => scope.includes("Sites.FullControl.All"))) {
      if (override.sharePointClaimsUnavailable) {
        return claimsUnavailable();
      }
      return ok({
        tid: override.sharePointTenantId ?? override.graphTenantId ?? tenantId,
        aud: override.sharePointAudience ?? tenantUrl,
      });
    }
    if (tokenRequest.scopes.some((scope) => MosServiceScope().includes(scope))) {
      if (override.mosClaimsUnavailable) {
        return claimsUnavailable();
      }
      return ok({
        tid: override.mosTenantId ?? tenantId,
        aud: override.mosAudience ?? "https://titles.example.test",
      });
    }
    return ok({
      tid: override.graphTenantId ?? tenantId,
      aud: override.graphAudience ?? "https://graph.microsoft.com",
    });
  });
}

/** Returns the exact bytes of one named part of a form-data request body. */
export function multipartPart(body: unknown, name: string): Buffer | undefined {
  if (!(body instanceof FormData)) {
    return undefined;
  }
  const raw = body.getBuffer();
  const header = raw.indexOf(`Content-Disposition: form-data; name="${name}"`);
  if (header < 0) {
    return undefined;
  }
  const start = raw.indexOf("\r\n\r\n", header) + 4;
  const end = raw.indexOf(`\r\n--${body.getBoundary()}`, start);
  return raw.subarray(start, end);
}
