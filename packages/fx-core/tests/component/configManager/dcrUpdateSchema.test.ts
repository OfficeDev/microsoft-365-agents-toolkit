// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { Ajv } from "ajv";
import { describe, expect, it } from "vitest";
import currentSchema from "../../../resource/yaml-schema/yaml.schema.json";
import versionedSchema from "../../../resource/yaml-schema/v1.13/yaml.schema.json";

const configurationId = "existing-dcr-configuration";
const fullInput = {
  configurationId,
  appId: "restricted-m365-app",
  resource: "https://resource.example.com/api",
  applicableToApps: "SpecificApp",
  targetAudience: "AnyTenant",
  includePersonalMicrosoftAccounts: true,
};

const validInputs: { title: string; input: Record<string, unknown> }[] = [
  { title: "configuration ID only", input: { configurationId } },
  { title: "all supported update fields", input: fullInput },
  {
    title: "enable MSA without a name",
    input: { configurationId, includePersonalMicrosoftAccounts: true },
  },
  {
    title: "disable MSA without a name",
    input: { configurationId, includePersonalMicrosoftAccounts: false },
  },
  { title: "resource only", input: { configurationId, resource: fullInput.resource } },
  {
    title: "app ID without resupplying the current mode",
    input: { configurationId, appId: fullInput.appId },
  },
  {
    title: "SpecificApp with an existing binding",
    input: { configurationId, applicableToApps: "SpecificApp" },
  },
  { title: "AnyApp", input: { configurationId, applicableToApps: "AnyApp" } },
  { title: "HomeTenant", input: { configurationId, targetAudience: "HomeTenant" } },
  { title: "AnyTenant", input: { configurationId, targetAudience: "AnyTenant" } },
];

const invalidInputs: { title: string; input: Record<string, unknown> }[] = [
  { title: "missing configuration ID", input: {} },
  { title: "empty configuration ID", input: { configurationId: "" } },
  { title: "non-string configuration ID", input: { configurationId: 42 } },
  { title: "empty name", input: { configurationId, name: "" } },
  { title: "overlong name", input: { configurationId, name: "n".repeat(129) } },
  { title: "immutable name", input: { configurationId, name: "Existing client" } },
  {
    title: "immutable MCP endpoint",
    input: { configurationId, mcpResourceUrl: "https://example.com/mcp" },
  },
  {
    title: "immutable metadata URL",
    input: {
      configurationId,
      wellKnownAuthorizationServer: "https://login.example.com/metadata",
    },
  },
  { title: "immutable empty targets", input: { configurationId, targetUrlsShouldStartWith: [] } },
  {
    title: "immutable multiple targets",
    input: {
      configurationId,
      targetUrlsShouldStartWith: ["https://one.example.com/mcp", "https://two.example.com/mcp"],
    },
  },
  {
    title: "supported update mixed with an immutable field",
    input: { ...fullInput, name: "New name" },
  },
  { title: "unsupported app mode", input: { configurationId, applicableToApps: "AllApps" } },
  { title: "unsupported tenant scope", input: { configurationId, targetAudience: "AllTenants" } },
  {
    title: "scalar targets",
    input: { configurationId, targetUrlsShouldStartWith: "https://example.com/mcp" },
  },
  { title: "non-string target entry", input: { configurationId, targetUrlsShouldStartWith: [42] } },
  {
    title: "string account option",
    input: { configurationId, includePersonalMicrosoftAccounts: "true" },
  },
  {
    title: "numeric account option",
    input: { configurationId, includePersonalMicrosoftAccounts: 1 },
  },
  { title: "static OAuth client ID", input: { configurationId, clientId: "static-client" } },
  { title: "static OAuth secret", input: { configurationId, clientSecret: "not-a-real-secret" } },
  { title: "OAuth base URL alias", input: { configurationId, baseUrl: "https://example.com/mcp" } },
  { title: "OpenAPI input", input: { configurationId, apiSpecPath: "./openapi.yaml" } },
  {
    title: "manual authorization endpoint",
    input: { configurationId, authorizationUrl: "https://login.example.com/authorize" },
  },
  {
    title: "manual token endpoint",
    input: { configurationId, tokenUrl: "https://login.example.com/token" },
  },
  {
    title: "manual refresh endpoint",
    input: { configurationId, refreshUrl: "https://login.example.com/refresh" },
  },
  { title: "manual scopes", input: { configurationId, scope: "read,write" } },
  { title: "PKCE input", input: { configurationId, isPKCEEnabled: true } },
  { title: "management policy input", input: { configurationId, manageableByUsers: ["user-id"] } },
  { title: "wire app ID alias", input: { configurationId, m365AppId: fullInput.appId } },
  {
    title: "wire account option alias",
    input: { configurationId, supportedAccountTypes: "Enterprise, Consumer" },
  },
  { title: "wire name alias", input: { configurationId, clientName: "Existing client" } },
  {
    title: "token exchange method",
    input: { configurationId, tokenExchangeMethodType: "PostRequestBody" },
  },
  { title: "identity provider conversion", input: { configurationId, identityProvider: "Custom" } },
  { title: "flow conversion", input: { configurationId, flow: "authorizationCode" } },
  { title: "unknown input", input: { configurationId, unknownSetting: true } },
];

for (const [field, value] of Object.entries(fullInput)) {
  invalidInputs.push({
    title: `null ${field}`,
    input: { configurationId, [field]: null },
  });
  if (typeof value === "string" && field !== "configurationId") {
    invalidInputs.push({
      title: `non-string ${field}`,
      input: { configurationId, [field]: 42 },
    });
  }
}

it("DCR-UPD-01: keeps both definitions identical and exposes exactly the supported update fields", () => {
  expect(currentSchema.definitions.dcrUpdate).toEqual(versionedSchema.definitions.dcrUpdate);
  expect(
    Object.keys(currentSchema.definitions.dcrUpdate.properties.with.properties).sort()
  ).toEqual(Object.keys(fullInput).sort());
  expect(currentSchema.definitions.dcrUpdate.properties.with.required).toEqual(["configurationId"]);
});

for (const { title, schema } of [
  { title: "current", schema: currentSchema },
  { title: "v1.13", schema: versionedSchema },
]) {
  describe(`dcr/update ${title} schema`, () => {
    const ajv = new Ajv({ allowUnionTypes: true });
    ajv.addKeyword("deprecationMessage");
    const validate = ajv.compile(schema);
    const document = (input: Record<string, unknown>) => ({
      version: "v1.13",
      provision: [{ uses: "dcr/update", with: input }],
    });

    for (const { title, input } of validInputs) {
      it(`DCR-UPD-01: accepts ${title}`, () => {
        expect(validate(document(input)), JSON.stringify(validate.errors)).toBe(true);
      });
    }

    for (const { title, input } of invalidInputs) {
      it(`DCR-UPD-02: rejects ${title}`, () => {
        expect(validate(document(input))).toBe(false);
      });
    }

    it("DCR-UPD-02: rejects an environment output mapping", () => {
      expect(
        validate({
          version: "v1.13",
          provision: [
            {
              uses: "dcr/update",
              with: { configurationId },
              writeToEnvironmentFile: { configurationId: "NEW_DCR_ID" },
            },
          ],
        })
      ).toBe(false);
    });

    it("DCR-UPD-24: keeps existing register and OAuth update YAML valid", () => {
      expect(
        validate({
          version: "v1.13",
          provision: [
            {
              uses: "dcr/register",
              with: {
                name: "Existing registration",
                mcpResourceUrl: "https://example.com/mcp",
              },
              writeToEnvironmentFile: { configurationId: "DCR_ID" },
            },
            {
              uses: "oauth/update",
              with: { configurationId, name: "Existing OAuth registration" },
            },
          ],
        }),
        JSON.stringify(validate.errors)
      ).toBe(true);
    });
  });
}
