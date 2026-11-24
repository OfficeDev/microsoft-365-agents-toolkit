// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { Ajv } from "ajv";
import { describe, expect, it } from "vitest";
import currentSchema from "../../../resource/yaml-schema/yaml.schema.json";
import versionedSchema from "../../../resource/yaml-schema/v1.14/yaml.schema.json";

const digest = "0123456789abcdef".repeat(4);
const hashInputNames = [
  "sourcePackageSha256",
  "sourceAppPackageSha256",
  "pluginSha256",
  "sharePointPackageSha256",
  "appPackageSha256",
];
const preparedInputs: Record<string, string> = {
  preparedPublicationPath: "${{PREP_PUBLICATION_PATH}}",
  publicationId: "${{PREP_PUBLICATION_ID}}",
  sourcePackageSha256: digest,
  sourceAppPackageSha256: "${{PREP_SOURCE_APP_PACKAGE_SHA256}}",
  pluginSha256: "${{PREP_PLUGIN_SHA256}}",
  sharePointPackageSha256: "${{PREP_SHAREPOINT_PACKAGE_SHA256}}",
  appPackageSha256: "${{PREP_APP_PACKAGE_SHA256}}",
};
const agentOutputs = { titleId: "M365_TITLE_ID", appId: "M365_APP_ID" };
const graphOutputs = { publishedAppId: "TEAMS_APP_PUBLISHED_APP_ID" };

for (const [name, schema] of [
  ["current", currentSchema],
  ["v1.14", versionedSchema],
] as const) {
  describe(`prepared publication ${name} schema`, () => {
    const ajv = new Ajv({ allowUnionTypes: true });
    ajv.addKeyword("deprecationMessage");
    const validate = ajv.compile(schema);
    const document = (
      uses: string,
      withInputs: Record<string, string>,
      outputs: Record<string, string>
    ) => ({
      version: "v1.14",
      publish: [{ uses, with: withInputs, writeToEnvironmentFile: outputs }],
    });
    const agentDocument = (withInputs: Record<string, string>) =>
      document("copilotAgent/publish", withInputs, agentOutputs);
    const graphDocument = (withInputs: Record<string, string>) =>
      document("teamsApp/publishAppPackage", withInputs, graphOutputs);

    it("SPFX-PREP-20: accepts the prepared inputs on copilotAgent/publish", () => {
      const withInputs = { appPackagePath: "./appPackage/build/app.zip", ...preparedInputs };

      expect(validate(agentDocument({ ...withInputs, scope: "tenant" })), "tenant").toBe(true);
      expect(validate(agentDocument(withInputs)), "default scope").toBe(true);
    });

    it("SPFX-PREP-20: keeps the unprepared copilotAgent/publish inputs valid", () => {
      expect(validate(agentDocument({ appPackagePath: "./app.zip", scope: "shared" }))).toBe(true);
    });

    for (const key of hashInputNames) {
      it(`SPFX-PREP-20: accepts a literal or environment-referenced ${key}`, () => {
        const base = { appPackagePath: "./app.zip", ...preparedInputs };

        expect(validate(agentDocument({ ...base, [key]: digest.toUpperCase() }))).toBe(true);
        expect(validate(agentDocument({ ...base, [key]: "${{PREPARED_HASH}}" }))).toBe(true);
      });

      it(`SPFX-PREP-20: rejects a malformed ${key}`, () => {
        const base = { appPackagePath: "./app.zip", ...preparedInputs };

        for (const malformed of ["abc", digest.slice(1), `${digest}0`, "g".repeat(64), "${{1}}"]) {
          expect(validate(agentDocument({ ...base, [key]: malformed })), malformed).toBe(false);
        }
      });
    }

    it("SPFX-PREP-20: rejects an unknown prepared input on copilotAgent/publish", () => {
      expect(
        validate(agentDocument({ appPackagePath: "./app.zip", ...preparedInputs, extra: "x" }))
      ).toBe(false);
    });

    it("SPFX-PREP-20: keeps teamsApp/publishAppPackage free of prepared inputs", () => {
      expect(validate(graphDocument({ appPackagePath: "./app.zip" }))).toBe(true);
      expect(validate(graphDocument({ appPackagePath: "./app.zip", ...preparedInputs }))).toBe(
        false
      );
      for (const [key, value] of Object.entries(preparedInputs)) {
        expect(validate(graphDocument({ appPackagePath: "./app.zip", [key]: value })), key).toBe(
          false
        );
      }
    });
  });
}
