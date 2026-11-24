// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { assert } from "chai";
import "mocha";
import sinon from "sinon";
import {
  AppManifestUtils,
  TeamsManifestConverter,
  TeamsManifestLatest,
  TeamsManifestWrapper,
} from "../src";
import * as fetchHelper from "../src/fetchHelper";

const schemaUrl =
  "https://developer.microsoft.com/json-schemas/teams/v1.31/MicrosoftTeams.schema.json";

function teamsManifest(): TeamsManifestLatest {
  return {
    $schema: schemaUrl,
    manifestVersion: "1.31",
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
  };
}

function activityManifest(displayName?: string) {
  return {
    ...teamsManifest(),
    activities: {
      activityTypes: [{ type: "task", description: "Task", templateText: "New task", displayName }],
    },
  };
}

function naaManifest(clientId?: string) {
  return {
    ...teamsManifest(),
    webApplicationInfo: {
      id: "aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee",
      nestedAppAuthInfo: [{ redirectUri: "https://example.com", scopes: ["User.Read"], clientId }],
    },
  };
}

function connectorManifest(type: string, referenceId?: string) {
  return {
    ...teamsManifest(),
    agentConnectors: [
      {
        id: "test",
        displayName: "Test",
        toolSource: {
          remoteMcpServer: {
            mcpServerUrl: "https://example.com/mcp",
            authorization: { type, referenceId },
          },
        },
      },
    ],
  };
}

async function validationErrors(input: unknown): Promise<string[]> {
  const manifest = TeamsManifestConverter.jsonToManifest(JSON.stringify(input));
  return AppManifestUtils.validateAgainstSchema(manifest);
}

describe("Teams app manifest v1.31", () => {
  it("validates with the packaged schema without a network request", async () => {
    const fetchStub = sinon.stub(fetchHelper, "default").rejects(new Error("Unexpected fetch"));
    try {
      assert.deepEqual(await validationErrors(teamsManifest()), []);
      sinon.assert.notCalled(fetchStub);
    } finally {
      fetchStub.restore();
    }
  });

  it("round-trips new fields through the registered converter and wrapper", async () => {
    const input = {
      ...activityManifest("New task"),
      webApplicationInfo: naaManifest("11111111-2222-3333-4444-555555555555").webApplicationInfo,
      agentConnectors: connectorManifest("ConnectionVault", "connection-reference").agentConnectors,
    };
    const manifest = TeamsManifestConverter.jsonToManifest(JSON.stringify(input));

    assert.deepEqual(JSON.parse(TeamsManifestConverter.manifestToJson(manifest)), input);
    assert.deepEqual(
      JSON.parse(TeamsManifestWrapper.fromJSON(JSON.stringify(input)).toJSON()),
      input
    );
    assert.deepEqual(await AppManifestUtils.validateAgainstSchema(manifest), []);
  });

  it("rejects malformed v1.31 input instead of falling back to unchecked conversion", () => {
    assert.throws(() => TeamsManifestConverter.jsonToManifest('{"manifestVersion":"1.31"}'));
    assert.throws(() =>
      TeamsManifestConverter.jsonToManifest(JSON.stringify(connectorManifest("UnsupportedAuth")))
    );
  });

  it("allows the new optional fields to be omitted", async () => {
    assert.deepEqual(await validationErrors(activityManifest()), []);
    assert.deepEqual(await validationErrors(naaManifest()), []);
  });

  it("validates activity displayName length and nonblank constraints", async () => {
    for (const displayName of ["A", "A".repeat(128)]) {
      assert.deepEqual(await validationErrors(activityManifest(displayName)), []);
    }
    for (const displayName of ["", " \t", "A".repeat(129)]) {
      const errors = await validationErrors(activityManifest(displayName));
      assert.isTrue(
        errors.some((error) => error.includes("/activities/activityTypes/0/displayName"))
      );
    }
  });

  it("requires a GUID for the nested-app clientId override", async () => {
    const errors = await validationErrors(naaManifest("not-a-guid"));
    assert.isTrue(
      errors.some((error) => error.includes("/webApplicationInfo/nestedAppAuthInfo/0/clientId"))
    );
  });

  it("requires referenceId for vault and dynamic-client authorization", async () => {
    for (const type of [
      "OAuthPluginVault",
      "ApiKeyPluginVault",
      "DynamicClientRegistration",
      "AzureKeyVault",
      "ConnectionVault",
    ]) {
      assert.deepEqual(await validationErrors(connectorManifest(type, "connection-reference")), []);
      const errors = await validationErrors(connectorManifest(type));
      assert.isTrue(
        errors.some((error) => error.includes("referenceId")),
        type
      );
    }
    assert.deepEqual(await validationErrors(connectorManifest("None")), []);
  });

  it("rejects removed contentBotId in v1.31 while retaining v1.30 support", async () => {
    const input = {
      ...teamsManifest(),
      staticTabs: [
        {
          entityId: "home",
          scopes: ["personal"],
          contentBotId: "aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee",
        },
      ],
    };
    assert.throws(
      () => TeamsManifestConverter.jsonToManifest(JSON.stringify(input)),
      "contentBotId"
    );
    const previous = {
      ...input,
      $schema: schemaUrl.replace("v1.31", "v1.30"),
      manifestVersion: "1.30",
    };
    assert.deepEqual(await validationErrors(previous), []);
  });
});
