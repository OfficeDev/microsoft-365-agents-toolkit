// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author yefuwang@microsoft.com
 */

import path from "path";
import { chai, vi } from "vitest";
import { YamlParser } from "../../../src/component/configManager/parser";
import fs from "fs-extra";
import Ajv from "ajv";

const assert: typeof chai.assert = chai.assert;

describe("v3 yaml parser", () => {
  describe("when parsing an invalid path", () => {
    before(() => {
      vi.spyOn(fs, "readFile").mockRejectedValue(new Error("file not found"));
    });

    afterEach(() => {
      vi.restoreAllMocks();
    });
    it("should return InvalidYamlSchemaError", async () => {
      const parser = new YamlParser();
      const result = await parser.parse("");
      assert(result.isErr() && result.error.name === "InvalidYamlSchemaError");
    });
  });

  describe("when parsing an empty file", () => {
    before(async () => {
      vi.spyOn(fs, "readFile").mockResolvedValue("");
    });

    after(() => {
      vi.restoreAllMocks();
    });

    it("should return InvalidYamlSchemaError", async () => {
      const parser = new YamlParser();
      const result = await parser.parse("");
      assert(result.isErr() && result.error.name === "InvalidYamlSchemaError");
    });
  });

  describe("when parsing a file containing only array", () => {
    it("should return InvalidYamlSchemaError", async () => {
      const parser = new YamlParser();
      const yamlPath = path.resolve(__dirname, "testing_data", "array.yml");
      const result = await parser.parse(yamlPath, true);
      assert(result.isErr() && result.error.name === "InvalidYamlSchemaError");
    });
  });

  describe("when parsing a file with lifecycle content not being array", () => {
    it("should return YamlFieldTypeError", async () => {
      const parser = new YamlParser();
      const result = await parser.parse(
        path.resolve(__dirname, "testing_data", "invalid_lifecycle_content.yml"),
        true
      );
      assert(result.isErr() && result.error.name === "InvalidYamlSchemaError");
    });
  });

  describe(`when parsing a file with lifecycle content with invalid "uses" and "with"`, () => {
    it("should return YamlFieldMissingError without 'with'", async () => {
      const parser = new YamlParser();
      const result = await parser.parse(
        path.resolve(__dirname, "testing_data", "invalid_lifecycle_without_with.yml"),
        true
      );
      assert(result.isErr() && result.error.name === "InvalidYamlSchemaError");
    });
    it("should return YamlFieldMissingError without 'uses'", async () => {
      const parser = new YamlParser();
      const result = await parser.parse(
        path.resolve(__dirname, "testing_data", "invalid_lifecycle_without_uses.yml"),
        true
      );
      assert(result.isErr() && result.error.name === "InvalidYamlSchemaError");
    });
    it("should return YamlFieldTypeError with wrong 'uses' type", async () => {
      const parser = new YamlParser();
      const result = await parser.parse(
        path.resolve(__dirname, "testing_data", "invalid_lifecycle_with_wrong_uses_type.yml"),
        true
      );
      assert(result.isErr() && result.error.name === "InvalidYamlSchemaError");
    });
    it("should return YamlFieldTypeError with wrong 'with' type", async () => {
      const parser = new YamlParser();
      const result = await parser.parse(
        path.resolve(__dirname, "testing_data", "invalid_lifecycle_with_wrong_with_type.yml"),
        true
      );
      assert(result.isErr() && result.error.name === "InvalidYamlSchemaError");
    });
  });

  describe(`when parsing a file with right schema, but unknown drivers`, () => {
    // because driver resolution happens when the driver actually runs.
    it("should return error", async () => {
      const parser = new YamlParser();
      const result = await parser.parse(
        path.resolve(__dirname, "testing_data", "valid_with_unknown_driver.yml"),
        true
      );
      assert(result.isErr() && result.error.name === "InvalidYamlSchemaError");
    });
  });

  describe(`when parsing real app.yml`, () => {
    // because driver resolution happens when the driver actually runs.
    it("should return ok", async () => {
      const parser = new YamlParser();
      const result = await parser.parse(path.resolve(__dirname, "testing_data", "app.yml"), true);
      assert(result.isOk());
      if (result.isOk()) {
        const model = result.value;
        chai.expect(model["provision"]).is.not.undefined;
        chai.expect(model["deploy"]).is.not.undefined;
        chai.expect(model["publish"]).is.not.undefined;
        chai.expect(model["configureApp"]).is.undefined;
        chai.expect(model["registerApp"]).is.undefined;
      }
    });
  });

  describe(`when parsing good_sample_tag.yml`, () => {
    it("should return ok", async () => {
      const parser = new YamlParser();
      const result = await parser.parse(
        path.resolve(__dirname, "testing_data", "good_sample_tag.yml"),
        true
      );
      assert(result.isOk());
      if (result.isOk()) {
        const model = result.value;
        chai.expect(model.additionalMetadata).is.not.undefined;
        chai.expect(model.additionalMetadata!["sampleTag"]).is.equal("testRepo:testSample");
      }
    });
  });

  describe(`when parsing bad_sample_tag.yml`, () => {
    it("should not return error", async () => {
      const parser = new YamlParser();
      const result = await parser.parse(
        path.resolve(__dirname, "testing_data", "bad_sample_tag.yml"),
        false
      );
      assert(result.isOk());
    });
  });

  describe(`when parsing yml with invalid env field`, () => {
    it("should return error if env field is of type string", async () => {
      const parser = new YamlParser();
      const result = await parser.parse(
        path.resolve(__dirname, "testing_data", "invalid_env_field_string.yml"),
        true
      );
      assert(result.isErr() && result.error.name === "InvalidYamlSchemaError");
    });

    it("should return error if env field value has wrong type", async () => {
      const parser = new YamlParser();
      const result = await parser.parse(
        path.resolve(__dirname, "testing_data", "invalid_env_subfield_type.yml"),
        true
      );
      assert(result.isErr() && result.error.name === "InvalidYamlSchemaError");
    });

    it("should return error if env field is of type array", async () => {
      const parser = new YamlParser();
      const result = await parser.parse(
        path.resolve(__dirname, "testing_data", "invalid_env_field_array.yml"),
        true
      );
      assert(result.isErr() && result.error.name === "InvalidYamlSchemaError");
    });
  });

  describe(`when parsing yml with valid env field`, async () => {
    it("should return ok", async () => {
      const parser = new YamlParser();
      const result = await parser.parse(
        path.resolve(__dirname, "testing_data", "valid_env_field.yml"),
        true
      );
      assert(result.isOk());
    });
  });

  describe(`when parsing yml with valid envrionmentFolderPath`, async () => {
    it("should return ok", async () => {
      const parser = new YamlParser();
      const result = await parser.parse(
        path.resolve(__dirname, "testing_data", "valid_env_folder_path.yml"),
        true
      );
      assert(result.isOk() && result.value.environmentFolderPath === "/home/xxx");
    });
  });

  describe(`when parsing yml with invalid folder path`, async () => {
    it("should return ok", async () => {
      const parser = new YamlParser();
      const result = await parser.parse(
        path.resolve(__dirname, "testing_data", "invalid_env_folder_path.yml"),
        true
      );
      assert(result.isErr() && result.error.name === "InvalidYamlSchemaError");
    });
  });

  describe(`when parsing yml with valid writeToEnvironmentFile`, async () => {
    it("should return ok", async () => {
      const parser = new YamlParser();
      const result = await parser.parse(
        path.resolve(__dirname, "testing_data", "valid_write_to_environment_file.yml"),
        true
      );
      assert(
        result.isOk() &&
          result.value.provision &&
          result.value.provision.driverDefs[0].writeToEnvironmentFile &&
          result.value.provision.driverDefs[0].writeToEnvironmentFile["botId"] === "XXX"
      );
    });
  });

  describe(`when parsing yml with dcr/register action`, () => {
    for (const schemaVersion of ["v1.13", "v1.14", ""]) {
      describe(`DCR-08: ${schemaVersion || "default"} schema`, () => {
        const schema = fs.readJSONSync(
          path.resolve(
            __dirname,
            "../../../resource/yaml-schema",
            schemaVersion,
            "yaml.schema.json"
          )
        );
        const ajv = new Ajv({ allowUnionTypes: true });
        ajv.addKeyword("deprecationMessage");
        const validate = ajv.compile(schema);
        const base = {
          name: "test-dcr",
          targetUrlsShouldStartWith: ["${{MCP_SERVER_URL}}"],
        };
        const legacy = { wellKnownAuthorizationServer: "${{AUTH_METADATA_URL}}" };
        const endpoint = { mcpResourceUrl: "${{MCP_SERVER_URL}}" };
        const validInputs = [
          { ...base, ...legacy },
          { ...base, ...endpoint },
          { ...base, ...legacy, ...endpoint, resource: "${{OAUTH_RESOURCE}}" },
        ];
        for (const input of validInputs) {
          it(`accepts ${Object.keys(input).join(", ")}`, () => {
            assert.isTrue(
              validate({
                version: schemaVersion || "v1.14",
                provision: [
                  {
                    uses: "dcr/register",
                    with: input,
                    writeToEnvironmentFile: { configurationId: "DCR_REGISTRATION_ID" },
                  },
                ],
              }),
              JSON.stringify(validate.errors)
            );
          });
        }
        const invalidInputs = [
          base,
          { name: "test-dcr", ...legacy },
          { ...base, ...legacy, targetUrlsShouldStartWith: [] },
          {
            ...base,
            ...legacy,
            targetUrlsShouldStartWith: ["https://one.example.com", "https://two.example.com"],
          },
          { ...base, ...legacy, targetUrlsShouldStartWith: "https://one.example.com" },
          { ...base, ...endpoint, authorizationServerUrl: "https://auth.example.com" },
          { ...base, ...endpoint, cimdSupported: false },
        ];
        for (const [index, input] of invalidInputs.entries()) {
          it(`rejects unsupported shape ${index + 1}`, () => {
            assert.isFalse(
              validate({
                version: schemaVersion || "v1.14",
                provision: [
                  {
                    uses: "dcr/register",
                    with: input,
                    writeToEnvironmentFile: { configurationId: "DCR_REGISTRATION_ID" },
                  },
                ],
              })
            );
          });
        }
      });
    }

    it("DCR-08: keeps default and versioned DCR definitions identical", () => {
      const root = path.resolve(__dirname, "../../../resource/yaml-schema");
      assert.deepEqual(
        fs.readJSONSync(path.join(root, "yaml.schema.json")).definitions.dcrRegister,
        fs.readJSONSync(path.join(root, "v1.14/yaml.schema.json")).definitions.dcrRegister
      );
    });

    it("DCR-08: parses endpoint-based YAML through the production parser", async () => {
      const input = {
        name: "test-dcr",
        mcpResourceUrl: "https://mcp.example.com/mcp",
        targetUrlsShouldStartWith: ["https://mcp.example.com/mcp"],
      };
      const read = vi.spyOn(fs, "readFile").mockResolvedValue(
        JSON.stringify({
          version: "v1.13",
          provision: [
            {
              uses: "dcr/register",
              with: input,
              writeToEnvironmentFile: { configurationId: "DCR_REGISTRATION_ID" },
            },
          ],
        })
      );
      try {
        const result = await new YamlParser().parse("m365agents.yml", true);
        assert.isTrue(result.isOk());
        if (result.isOk()) {
          assert.deepEqual(result.value.provision?.driverDefs[0].with, input);
        }
      } finally {
        read.mockRestore();
      }
    });

    for (const supportedAccountTypes of [
      "Enterprise",
      "Enterprise, Consumer",
      "Consumer, Enterprise",
    ]) {
      it(`OAUTH-MSA-02: v1.14 accepts ${supportedAccountTypes} for OAuth and DCR actions`, () => {
        const schema = fs.readJSONSync(
          path.resolve(__dirname, "../../../resource/yaml-schema/v1.14/yaml.schema.json")
        );
        const ajv = new Ajv({ allowUnionTypes: true });
        ajv.addKeyword("deprecationMessage");
        const validate = ajv.compile(schema);

        assert.isTrue(
          validate({
            version: "v1.14",
            provision: [
              {
                uses: "oauth/register",
                with: { name: "test-oauth", flow: "authorizationCode", supportedAccountTypes },
                writeToEnvironmentFile: { configurationId: "OAUTH_CONFIGURATION_ID" },
              },
              {
                uses: "oauth/update",
                with: {
                  name: "test-oauth",
                  configurationId: "existing-id",
                  supportedAccountTypes,
                },
              },
              {
                uses: "dcr/register",
                with: {
                  name: "test-dcr",
                  mcpResourceUrl: "https://mcp.example.com/mcp",
                  targetUrlsShouldStartWith: ["https://mcp.example.com/mcp"],
                  supportedAccountTypes,
                },
                writeToEnvironmentFile: { configurationId: "DCR_CONFIGURATION_ID" },
              },
            ],
          }),
          JSON.stringify(validate.errors)
        );
      });
    }

    it("OAUTH-MSA-03: v1.14 rejects consumer-only account types", () => {
      const schema = fs.readJSONSync(
        path.resolve(__dirname, "../../../resource/yaml-schema/v1.14/yaml.schema.json")
      );
      const ajv = new Ajv({ allowUnionTypes: true });
      ajv.addKeyword("deprecationMessage");
      const validate = ajv.compile(schema);

      assert.isFalse(
        validate({
          version: "v1.14",
          provision: [
            {
              uses: "oauth/register",
              with: {
                name: "test-oauth",
                flow: "authorizationCode",
                supportedAccountTypes: "Consumer",
              },
              writeToEnvironmentFile: { configurationId: "OAUTH_CONFIGURATION_ID" },
            },
          ],
        })
      );
    });

    it("should return ok for valid dcr/register", async () => {
      const parser = new YamlParser();
      const result = await parser.parse(
        path.resolve(__dirname, "testing_data", "valid_dcr_register.yml"),
        true
      );
      assert(result.isOk());
      if (result.isOk()) {
        const def = result.value.provision?.driverDefs[0];
        chai.expect(def?.uses).to.equal("dcr/register");
        chai
          .expect((def?.with as any)?.wellKnownAuthorizationServer)
          .to.equal("https://example.com/.well-known/oauth-authorization-server");
        chai.expect(def?.writeToEnvironmentFile?.configurationId).to.equal("DCR_REGISTRATION_ID");
      }
    });

    it("should return InvalidYamlSchemaError when wellKnownAuthorizationServer is missing", async () => {
      const parser = new YamlParser();
      const result = await parser.parse(
        path.resolve(__dirname, "testing_data", "invalid_dcr_register_missing_well_known.yml"),
        true
      );
      assert(result.isErr() && result.error.name === "InvalidYamlSchemaError");
    });
  });

  describe(`when parsing yml with invalid writeToEnvironmentFile`, async () => {
    it("should return YamlFieldTypeError", async () => {
      const parser = new YamlParser();
      let result = await parser.parse(
        path.resolve(
          __dirname,
          "testing_data",
          "invalid_write_to_environment_file_array_teamsapp.yml"
        ),
        true
      );
      assert(result.isErr() && result.error.name === "InvalidYamlSchemaError");
      const errorMsg = result._unsafeUnwrapErr().message;
      chai
        .expect(errorMsg)
        .includes(`Unable to parse yaml file`)
        .and.includes(`Please open the yaml file`);

      result = await parser.parse(
        path.resolve(__dirname, "testing_data", "invalid_write_to_environment_file_number.yml"),
        true
      );
      assert(result.isErr() && result.error.name === "InvalidYamlSchemaError");
    });
  });
});
