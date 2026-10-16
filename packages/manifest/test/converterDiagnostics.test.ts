// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { assert } from "chai";
import "mocha";
import {
  ApiPluginManifestConverter,
  DeclarativeAgentManifestConverter,
  TeamsManifestConverter,
} from "../src";

function manifestWithConnectors(agentConnectors: unknown): string {
  return JSON.stringify({
    manifestVersion: "devPreview",
    version: "1.0.0",
    id: "aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee",
    developer: {
      name: "Test",
      websiteUrl: "https://example.com",
      privacyUrl: "https://example.com/privacy",
      termsOfUseUrl: "https://example.com/terms",
    },
    name: { short: "Test", full: "Test app" },
    description: { short: "Test", full: "Test app" },
    icons: { color: "color.png", outline: "outline.png" },
    accentColor: "#FFFFFF",
    agentConnectors,
  });
}

function remoteConnector(type: string) {
  return {
    id: "test",
    displayName: "Test",
    toolSource: {
      remoteMcpServer: {
        mcpServerUrl: "https://example.com/mcp",
        authorization: { type, referenceId: "private-reference" },
      },
    },
  };
}

function conversionError(
  json: string,
  convert: (json: string) => unknown = TeamsManifestConverter.jsonToManifest
): Error {
  try {
    convert(json);
  } catch (error) {
    if (error instanceof Error) {
      return error;
    }
    throw error;
  }
  return assert.fail("Expected a converter error");
}

describe("Converter diagnostics", () => {
  it("reports the nested authorization enum instead of blaming the connector array", () => {
    const error = conversionError(manifestWithConnectors([remoteConnector("UnsupportedAuth")]));

    assert.include(
      error.message,
      "agentConnectors[0].toolSource.remoteMcpServer.authorization.type"
    );
    assert.include(error.message, "OAuthPluginVault");
    assert.notInclude(error.message, "optional object");
    assert.notInclude(error.message, "private-reference");
    assert.notInclude(error.message, "https://example.com/mcp");
  });

  it("reports the index of the invalid element after a valid connector", () => {
    const error = conversionError(
      manifestWithConnectors([remoteConnector("None"), remoteConnector("UnsupportedAuth")])
    );

    assert.include(
      error.message,
      "agentConnectors[1].toolSource.remoteMcpServer.authorization.type"
    );
  });

  it("identifies unsupported properties without printing their contents", () => {
    const error = conversionError(
      manifestWithConnectors([
        {
          id: "test",
          displayName: "Test",
          toolSource: {
            unsupportedConnector: {
              authorization: { type: "ConnectionVault", referenceId: "private-reference" },
            },
          },
        },
      ])
    );

    assert.include(error.message, "agentConnectors[0].toolSource.unsupportedConnector");
    assert.include(error.message, "not supported");
    assert.notInclude(error.message, "private-reference");
    assert.notInclude(error.message, "ConnectionVault");
  });

  it("describes an optional array correctly when the input is an object", () => {
    const error = conversionError(manifestWithConnectors({}));

    assert.include(error.message, "agentConnectors");
    assert.include(error.message, "array");
    assert.include(error.message, "got object");
    assert.notInclude(error.message, "optional object");
  });

  it("preserves missing required property errors inside optional arrays", () => {
    const error = conversionError(manifestWithConnectors([{ id: "test" }]));

    assert.include(error.message, "agentConnectors[0].displayName");
    assert.include(error.message, "required");
    assert.include(error.message, "string");
  });

  it("describes null array elements as null rather than object", () => {
    const error = conversionError(manifestWithConnectors([null]));

    assert.include(error.message, "agentConnectors[0]");
    assert.include(error.message, "got null");
  });

  it("does not serialize invalid scalar input", () => {
    const error = conversionError(
      manifestWithConnectors([{ id: { secret: "private-value" }, displayName: "Test" }])
    );

    assert.include(error.message, "agentConnectors[0].id");
    assert.include(error.message, "string");
    assert.include(error.message, "got object");
    assert.notInclude(error.message, "private-value");
  });

  it("preserves JSON syntax errors", () => {
    assert.throws(() => TeamsManifestConverter.jsonToManifest("{"), SyntaxError);
  });

  it("uses the same safe diagnostics for declarative agent converters", () => {
    const error = conversionError(
      JSON.stringify({ version: "v1.8", name: { secret: "private-value" } }),
      DeclarativeAgentManifestConverter.jsonToManifest
    );

    assert.include(error.message, '"name"');
    assert.include(error.message, "string");
    assert.include(error.message, "got object");
    assert.notInclude(error.message, "private-value");
  });

  it("uses the same safe diagnostics for API plugin converters", () => {
    const error = conversionError(
      JSON.stringify({ schema_version: "v2.4", name_for_human: { secret: "private-value" } }),
      ApiPluginManifestConverter.jsonToManifest
    );

    assert.include(error.message, '"name_for_human"');
    assert.include(error.message, "string");
    assert.include(error.message, "got object");
    assert.notInclude(error.message, "private-value");
  });

  for (const connectors of [
    undefined,
    [],
    [remoteConnector("None")],
    [remoteConnector("OAuthPluginVault")],
    [remoteConnector("ApiKeyPluginVault")],
    [remoteConnector("DynamicClientRegistration")],
    [remoteConnector("AzureKeyVault")],
    [remoteConnector("ConnectionVault")],
  ]) {
    it(`preserves valid connector round trips: ${JSON.stringify(connectors)}`, () => {
      const json = manifestWithConnectors(connectors);
      const manifest = TeamsManifestConverter.jsonToManifest(json);
      const roundTrip = TeamsManifestConverter.manifestToJson(manifest);

      assert.deepEqual(JSON.parse(roundTrip), JSON.parse(json));
    });
  }
});
