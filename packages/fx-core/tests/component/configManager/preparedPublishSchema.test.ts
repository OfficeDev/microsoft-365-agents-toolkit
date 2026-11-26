// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { Ajv } from "ajv";
import { describe, expect, it } from "vitest";
import currentSchema from "../../../resource/yaml-schema/yaml.schema.json";
import versionedSchema from "../../../resource/yaml-schema/v1.14/yaml.schema.json";

const outputNames = {
  sharePointPackagePath: "PREPARED_SPPKG",
  appPackagePath: "PREPARED_APP_PACKAGE",
};

const publishLifecycle = [
  {
    uses: "spfx/preparePublication",
    with: {
      packagePath: "./SPFx/sharepoint/solution.sppkg",
      appPackagePath: "./appPackage/build/app.zip",
    },
    writeToEnvironmentFile: outputNames,
  },
  {
    uses: "spfx/deploy",
    with: {
      packageSolutionPath: "./SPFx/config/package-solution.json",
      packagePath: "${{PREPARED_SPPKG}}",
    },
  },
  {
    uses: "copilotAgent/publish",
    with: {
      appPackagePath: "${{PREPARED_APP_PACKAGE}}",
    },
    writeToEnvironmentFile: {
      titleId: "M365_TITLE_ID",
      appId: "M365_APP_ID",
    },
  },
];

for (const [name, schema] of [
  ["current", currentSchema],
  ["v1.14", versionedSchema],
] as const) {
  describe(`prepared publication ${name} schema`, () => {
    const ajv = new Ajv({ allowUnionTypes: true });
    ajv.addKeyword("deprecationMessage");
    const validate = ajv.compile(schema);

    it("SPFX-PREP-11: accepts preparation outputs wired to ordinary deploy and publish actions", () => {
      expect(validate({ version: "v1.14", publish: publishLifecycle })).toBe(true);
    });

    it("SPFX-PREP-11: accepts ordinary deploy and publish actions without preparation outputs", () => {
      expect(
        validate({
          version: "v1.14",
          publish: [
            {
              uses: "spfx/deploy",
              with: { packageSolutionPath: "./SPFx/config/package-solution.json" },
            },
            {
              uses: "copilotAgent/publish",
              with: { appPackagePath: "./appPackage/build/app.zip" },
              writeToEnvironmentFile: {
                titleId: "M365_TITLE_ID",
                appId: "M365_APP_ID",
              },
            },
          ],
        })
      ).toBe(true);
    });

    it("SPFX-PREP-11: rejects removed prepared-mode inputs on existing actions", () => {
      expect(
        validate({
          version: "v1.14",
          publish: [
            {
              uses: "spfx/deploy",
              with: {
                packageSolutionPath: "./SPFx/config/package-solution.json",
                preparedPublicationPath: "${{PREP_PUBLICATION_PATH}}",
              },
            },
          ],
        })
      ).toBe(false);
      expect(
        validate({
          version: "v1.14",
          publish: [
            {
              uses: "copilotAgent/publish",
              with: {
                appPackagePath: "./appPackage/build/app.zip",
                publicationId: "${{PREP_PUBLICATION_ID}}",
              },
            },
          ],
        })
      ).toBe(false);
    });
  });
}
