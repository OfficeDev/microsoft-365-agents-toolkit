// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { assert } from "chai";
import "mocha";
import fs from "fs-extra";
import path from "path";
import sinon from "sinon";
import {
  AppManifestUtils,
  TeamsManifestConverter,
  TeamsManifestLatest,
  TeamsManifestV1D31,
  TeamsManifestWrapper,
} from "../src";
import { quicktypeJSONSchema } from "../convert";
import * as fetchHelper from "../src/fetchHelper";

const schemaUrl =
  "https://developer.microsoft.com/en-us/json-schemas/teams/v1.31/MicrosoftTeams.schema.json";
const appId = "11111111-1111-4111-8111-111111111111";

function teamsManifest() {
  return {
    $schema: schemaUrl,
    manifestVersion: "1.31",
    version: "1.0.0",
    id: appId,
    developer: {
      name: "Test developer",
      websiteUrl: "https://example.com",
      privacyUrl: "https://example.com/privacy",
      termsOfUseUrl: "https://example.com/terms",
    },
    icons: { color: "color.png", outline: "outline.png" },
    name: { short: "Test app" },
    description: { short: "Test app", full: "Test app description" },
    accentColor: "#FFFFFF",
    webApplicationInfo: {
      id: appId,
      nestedAppAuthInfo: [
        { redirectUri: "https://example.com", scopes: ["User.Read"], clientId: appId },
      ],
    },
    activities: {
      activityTypes: [
        {
          type: "task",
          description: "Task notification",
          templateText: "New task",
          displayName: "Task",
        },
      ],
    },
    agentConnectors: [
      {
        id: "connector",
        displayName: "Test connector",
        toolSource: {
          remoteMcpServer: {
            mcpServerUrl: "https://example.com/mcp",
            authorization: { type: "ConnectionVault", referenceId: "connection-reference" },
          },
        },
      },
    ],
  } satisfies TeamsManifestLatest;
}

describe("Teams manifest v1.31", () => {
  it("round-trips the new stable fields through the registered converter and wrapper", async () => {
    const input = teamsManifest();
    const manifest = TeamsManifestConverter.jsonToManifest(JSON.stringify(input));

    assert.deepEqual(JSON.parse(TeamsManifestConverter.manifestToJson(manifest)), input);
    assert.deepEqual(JSON.parse(TeamsManifestV1D31.Convert.teamsManifestV1D31ToJson(input)), input);
    assert.deepEqual(await AppManifestUtils.validateAgainstSchema(manifest), []);
    const wrapper = TeamsManifestWrapper.fromJSON(JSON.stringify(input));
    assert.deepEqual(JSON.parse(wrapper.toJSON()), input);
  });

  it("keeps the v1.31 model and converter in sync with the upstream snapshot", async () => {
    const root = path.join(__dirname, "..", "src");
    const schema = await fs.readFile(
      path.join(root, "json-schemas", "teams", "v1.31", "MicrosoftTeams.schema.json"),
      "utf8"
    );
    const { lines } = await quicktypeJSONSchema("typescript", "TeamsManifestV1D31", schema);
    const source = await fs.readFile(
      path.join(root, "generated-types", "teams", "TeamsManifestV1D31.ts"),
      "utf8"
    );
    const copyright =
      "// Copyright (c) Microsoft Corporation.\n// Licensed under the MIT license.\n\n";
    assert.strictEqual(source.replace(copyright, ""), lines.join("\n"));
  });

  it("rejects malformed v1.31 fields instead of using the unchecked fallback", () => {
    assert.throws(() =>
      TeamsManifestConverter.jsonToManifest(JSON.stringify({ manifestVersion: "1.31" }))
    );
    const input = teamsManifest();
    assert.throws(
      () =>
        TeamsManifestConverter.jsonToManifest(
          JSON.stringify({
            ...input,
            activities: {
              activityTypes: [{ ...input.activities.activityTypes[0], displayName: 42 }],
            },
          })
        ),
      /activities\.activityTypes\[0\]\.displayName/
    );
  });

  it("loads the primary and companion v1.31 snapshots without a network request", async () => {
    const fetchStub = sinon.stub(fetchHelper, "default").rejects(new Error("Unexpected fetch"));
    try {
      for (const name of ["", "AgenticUser.", "Localization.", "ResponseRenderingTemplate."]) {
        const schema = await AppManifestUtils.fetchSchema(
          schemaUrl.replace("MicrosoftTeams.schema.json", `MicrosoftTeams.${name}schema.json`)
        );
        assert.isObject(schema);
      }
      assert.isTrue(fetchStub.notCalled);
    } finally {
      fetchStub.restore();
    }
  });

  it("enforces clientId, displayName, and ConnectionVault schema constraints", async () => {
    const badClientId = teamsManifest();
    badClientId.webApplicationInfo.nestedAppAuthInfo[0].clientId = "not-a-guid";
    const emptyDisplayName = teamsManifest();
    emptyDisplayName.activities.activityTypes[0].displayName = "";
    const missingReference = teamsManifest();
    const authorization =
      missingReference.agentConnectors[0].toolSource.remoteMcpServer.authorization;
    const withoutReference = { type: authorization.type };
    const missingReferenceInput = {
      ...missingReference,
      agentConnectors: [
        {
          ...missingReference.agentConnectors[0],
          toolSource: {
            remoteMcpServer: {
              ...missingReference.agentConnectors[0].toolSource.remoteMcpServer,
              authorization: withoutReference,
            },
          },
        },
      ],
    };

    for (const [input, errorPath] of [
      [badClientId, "/webApplicationInfo/nestedAppAuthInfo/0/clientId"],
      [emptyDisplayName, "/activities/activityTypes/0/displayName"],
      [missingReferenceInput, "referenceId"],
    ] as const) {
      const manifest = TeamsManifestConverter.jsonToManifest(JSON.stringify(input));
      const errors = await AppManifestUtils.validateAgainstSchema(manifest);
      assert.isTrue(
        errors.some((error) => error.includes(errorPath)),
        errors.join("\n")
      );
    }
  });

  it("keeps the v1.30 converter available and rejects the removed v1.31 contentBotId", () => {
    const staticTabs = [{ entityId: "home", scopes: ["personal"], contentBotId: appId }];
    const previous = {
      ...teamsManifest(),
      webApplicationInfo: undefined,
      activities: undefined,
      agentConnectors: undefined,
      $schema: schemaUrl.replace("v1.31", "v1.30"),
      manifestVersion: "1.30",
      staticTabs,
    };
    const manifest = TeamsManifestConverter.jsonToManifest(JSON.stringify(previous));
    assert.deepEqual(
      JSON.parse(TeamsManifestConverter.manifestToJson(manifest)),
      JSON.parse(JSON.stringify(previous))
    );
    assert.throws(() =>
      TeamsManifestConverter.jsonToManifest(
        JSON.stringify({ ...previous, $schema: schemaUrl, manifestVersion: "1.31" })
      )
    );
  });
});
