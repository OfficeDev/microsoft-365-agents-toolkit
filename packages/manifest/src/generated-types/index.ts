// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import Ajv, { JSONSchemaType } from "ajv";
import addFormats from "ajv-formats";
import Ajv2020 from "ajv/dist/2020";
import fs from "fs-extra";
import {
  CopilotDeclarativeAgentV1D0,
  Convert as CopilotDeclarativeAgentV1D0Convert,
} from "./copilot/declarative-agent/CopilotDeclarativeAgent.v1d0";
import {
  CopilotDeclarativeAgentV1D2,
  Convert as CopilotDeclarativeAgentV1D2Convert,
} from "./copilot/declarative-agent/CopilotDeclarativeAgent.v1d2";
import {
  CopilotDeclarativeAgentV1D3,
  Convert as CopilotDeclarativeAgentV1D3Convert,
} from "./copilot/declarative-agent/CopilotDeclarativeAgent.v1d3";
import {
  CopilotPluginV2D1,
  Convert as CopilotPluginV2D1Convert,
} from "./copilot/plugin/CopilotPlugin.v2d1";
import {
  CopilotPluginV2D2,
  Convert as CopilotPluginV2D2Convert,
} from "./copilot/plugin/CopilotPlugin.v2d2";
import {
  MicrosoftTeamsV1D1, Convert as MicrosoftTeamsV1D1Convert,
} from "./teams/MicrosoftTeams.v1d1";
import {
  MicrosoftTeamsV1D10,
  Convert as MicrosoftTeamsV1D10Convert,
} from "./teams/MicrosoftTeams.v1d10";
import {
  MicrosoftTeamsV1D11,
  Convert as MicrosoftTeamsV1D11Convert,
} from "./teams/MicrosoftTeams.v1d11";
import {
  MicrosoftTeamsV1D12,
  Convert as MicrosoftTeamsV1D12Convert,
} from "./teams/MicrosoftTeams.v1d12";
import {
  MicrosoftTeamsV1D13,
  Convert as MicrosoftTeamsV1D13Convert,
} from "./teams/MicrosoftTeams.v1d13";
import {
  MicrosoftTeamsV1D14,
  Convert as MicrosoftTeamsV1D14Convert,
} from "./teams/MicrosoftTeams.v1d14";
import {
  MicrosoftTeamsV1D15,
  Convert as MicrosoftTeamsV1D15Convert,
} from "./teams/MicrosoftTeams.v1d15";
import {
  MicrosoftTeamsV1D16,
  Convert as MicrosoftTeamsV1D16Convert,
} from "./teams/MicrosoftTeams.v1d16";
import {
  MicrosoftTeamsV1D17,
  Convert as MicrosoftTeamsV1D17Convert,
} from "./teams/MicrosoftTeams.v1d17";
import {
  MicrosoftTeamsV1D19,
  Convert as MicrosoftTeamsV1D19Convert,
} from "./teams/MicrosoftTeams.v1d19";
import {
  MicrosoftTeamsV1D2,
  Convert as MicrosoftTeamsV1D2Convert,
} from "./teams/MicrosoftTeams.v1d2";
import {
  MicrosoftTeamsV1D20,
  Convert as MicrosoftTeamsV1D20Convert,
} from "./teams/MicrosoftTeams.v1d20";
import {
  MicrosoftTeamsV1D3,
  Convert as MicrosoftTeamsV1D3Convert,
} from "./teams/MicrosoftTeams.v1d3";
import {
  MicrosoftTeamsV1D4,
  Convert as MicrosoftTeamsV1D4Convert,
} from "./teams/MicrosoftTeams.v1d4";
import {
  MicrosoftTeamsV1D5,
  Convert as MicrosoftTeamsV1D5Convert,
} from "./teams/MicrosoftTeams.v1d5";
import {
  MicrosoftTeamsV1D6,
  Convert as MicrosoftTeamsV1D6Convert,
} from "./teams/MicrosoftTeams.v1d6";
import {
  MicrosoftTeamsV1D7,
  Convert as MicrosoftTeamsV1D7Convert,
} from "./teams/MicrosoftTeams.v1d7";
import {
  MicrosoftTeamsV1D8,
  Convert as MicrosoftTeamsV1D8Convert,
} from "./teams/MicrosoftTeams.v1d8";
import {
  MicrosoftTeamsV1D9,
  Convert as MicrosoftTeamsV1D9Convert,
} from "./teams/MicrosoftTeams.v1d9";
import {
  MicrosoftTeamsVDevPreview,
  Convert as MicrosoftTeamsVDevPreviewConvert,
} from "./teams/MicrosoftTeams.vDevPreview";

export {
  MicrosoftTeamsV1D1,
  Convert as MicrosoftTeamsV1D1Convert
} from "./teams/MicrosoftTeams.v1d1";
export {
  MicrosoftTeamsV1D10,
  Convert as MicrosoftTeamsV1D10Convert
} from "./teams/MicrosoftTeams.v1d10";
export {
  MicrosoftTeamsV1D11,
  Convert as MicrosoftTeamsV1D11Convert
} from "./teams/MicrosoftTeams.v1d11";
export {
  MicrosoftTeamsV1D12,
  Convert as MicrosoftTeamsV1D12Convert
} from "./teams/MicrosoftTeams.v1d12";
export {
  MicrosoftTeamsV1D13,
  Convert as MicrosoftTeamsV1D13Convert
} from "./teams/MicrosoftTeams.v1d13";
export {
  MicrosoftTeamsV1D14,
  Convert as MicrosoftTeamsV1D14Convert
} from "./teams/MicrosoftTeams.v1d14";
export {
  MicrosoftTeamsV1D15,
  Convert as MicrosoftTeamsV1D15Convert
} from "./teams/MicrosoftTeams.v1d15";
export {
  MicrosoftTeamsV1D16,
  Convert as MicrosoftTeamsV1D16Convert
} from "./teams/MicrosoftTeams.v1d16";
export {
  MicrosoftTeamsV1D17,
  Convert as MicrosoftTeamsV1D17Convert
} from "./teams/MicrosoftTeams.v1d17";
export {
  MicrosoftTeamsV1D19,
  Convert as MicrosoftTeamsV1D19Convert
} from "./teams/MicrosoftTeams.v1d19";
export {
  MicrosoftTeamsV1D2,
  Convert as MicrosoftTeamsV1D2Convert
} from "./teams/MicrosoftTeams.v1d2";
export {
  MicrosoftTeamsV1D20,
  Convert as MicrosoftTeamsV1D20Convert
} from "./teams/MicrosoftTeams.v1d20";
export {
  MicrosoftTeamsV1D3,
  Convert as MicrosoftTeamsV1D3Convert
} from "./teams/MicrosoftTeams.v1d3";
export {
  MicrosoftTeamsV1D4,
  Convert as MicrosoftTeamsV1D4Convert
} from "./teams/MicrosoftTeams.v1d4";
export {
  MicrosoftTeamsV1D5,
  Convert as MicrosoftTeamsV1D5Convert
} from "./teams/MicrosoftTeams.v1d5";
export {
  MicrosoftTeamsV1D6,
  Convert as MicrosoftTeamsV1D6Convert
} from "./teams/MicrosoftTeams.v1d6";
export {
  MicrosoftTeamsV1D7,
  Convert as MicrosoftTeamsV1D7Convert
} from "./teams/MicrosoftTeams.v1d7";
export {
  MicrosoftTeamsV1D8,
  Convert as MicrosoftTeamsV1D8Convert
} from "./teams/MicrosoftTeams.v1d8";
export {
  MicrosoftTeamsV1D9,
  Convert as MicrosoftTeamsV1D9Convert
} from "./teams/MicrosoftTeams.v1d9";
export { MicrosoftTeamsVDevPreview as DevPreviewSchema, MicrosoftTeamsVDevPreview, Convert as MicrosoftTeamsVDevPreviewConvert } from "./teams/MicrosoftTeams.vDevPreview";
export type MicrosoftTeamsManifest =
  | MicrosoftTeamsV1D1
  | MicrosoftTeamsV1D2
  | MicrosoftTeamsV1D3
  | MicrosoftTeamsV1D4
  | MicrosoftTeamsV1D5
  | MicrosoftTeamsV1D6
  | MicrosoftTeamsV1D7
  | MicrosoftTeamsV1D8
  | MicrosoftTeamsV1D9
  | MicrosoftTeamsV1D10
  | MicrosoftTeamsV1D11
  | MicrosoftTeamsV1D12
  | MicrosoftTeamsV1D13
  | MicrosoftTeamsV1D14
  | MicrosoftTeamsV1D15
  | MicrosoftTeamsV1D16
  | MicrosoftTeamsV1D17
  | MicrosoftTeamsV1D19
  | MicrosoftTeamsV1D20
  | MicrosoftTeamsVDevPreview;

export {
  CopilotDeclarativeAgentV1D0,
  Convert as CopilotDeclarativeAgentV1D0Convert
} from "./copilot/declarative-agent/CopilotDeclarativeAgent.v1d0";
export {
  CopilotDeclarativeAgentV1D2,
  Convert as CopilotDeclarativeAgentV1D2Convert
} from "./copilot/declarative-agent/CopilotDeclarativeAgent.v1d2";
export {
  CopilotDeclarativeAgentV1D3,
  Convert as CopilotDeclarativeAgentV1D3Convert
} from "./copilot/declarative-agent/CopilotDeclarativeAgent.v1d3";
export type DeclarativeAgentManifest =
  | CopilotDeclarativeAgentV1D0
  | CopilotDeclarativeAgentV1D2
  | CopilotDeclarativeAgentV1D3;

export {
  CopilotPluginV2D1,
  Convert as CopilotPluginV2D1Convert
} from "./copilot/plugin/CopilotPlugin.v2d1";
export {
  CopilotPluginV2D2,
  Convert as CopilotPluginV2D2Convert
} from "./copilot/plugin/CopilotPlugin.v2d2";
export type ApiPluginManifest = CopilotPluginV2D1 | CopilotPluginV2D2;

export type AppManifest = MicrosoftTeamsManifest | DeclarativeAgentManifest | ApiPluginManifest;

type Converters = {
  [key: string]: [(json: string) => any, (manifest: any) => string];
};
const TeamsManifestConverterMap: Converters = {
  "1.1": [
    MicrosoftTeamsV1D1Convert.toMicrosoftTeamsV1D1,
    MicrosoftTeamsV1D1Convert.microsoftTeamsV1D1ToJson,
  ],
  "1.2": [
    MicrosoftTeamsV1D2Convert.toMicrosoftTeamsV1D2,
    MicrosoftTeamsV1D2Convert.microsoftTeamsV1D2ToJson,
  ],
  "1.3": [
    MicrosoftTeamsV1D3Convert.toMicrosoftTeamsV1D3,
    MicrosoftTeamsV1D3Convert.microsoftTeamsV1D3ToJson,
  ],
  "1.4": [
    MicrosoftTeamsV1D4Convert.toMicrosoftTeamsV1D4,
    MicrosoftTeamsV1D4Convert.microsoftTeamsV1D4ToJson,
  ],
  "1.5": [
    MicrosoftTeamsV1D5Convert.toMicrosoftTeamsV1D5,
    MicrosoftTeamsV1D5Convert.microsoftTeamsV1D5ToJson,
  ],
  "1.6": [
    MicrosoftTeamsV1D6Convert.toMicrosoftTeamsV1D6,
    MicrosoftTeamsV1D6Convert.microsoftTeamsV1D6ToJson,
  ],
  "1.7": [
    MicrosoftTeamsV1D7Convert.toMicrosoftTeamsV1D7,
    MicrosoftTeamsV1D7Convert.microsoftTeamsV1D7ToJson,
  ],
  "1.8": [
    MicrosoftTeamsV1D8Convert.toMicrosoftTeamsV1D8,
    MicrosoftTeamsV1D8Convert.microsoftTeamsV1D8ToJson,
  ],
  "1.9": [
    MicrosoftTeamsV1D9Convert.toMicrosoftTeamsV1D9,
    MicrosoftTeamsV1D9Convert.microsoftTeamsV1D9ToJson,
  ],
  "1.10": [
    MicrosoftTeamsV1D10Convert.toMicrosoftTeamsV1D10,
    MicrosoftTeamsV1D10Convert.microsoftTeamsV1D10ToJson,
  ],
  "1.11": [
    MicrosoftTeamsV1D11Convert.toMicrosoftTeamsV1D11,
    MicrosoftTeamsV1D11Convert.microsoftTeamsV1D11ToJson,
  ],
  "1.12": [
    MicrosoftTeamsV1D12Convert.toMicrosoftTeamsV1D12,
    MicrosoftTeamsV1D12Convert.microsoftTeamsV1D12ToJson,
  ],
  "1.13": [
    MicrosoftTeamsV1D13Convert.toMicrosoftTeamsV1D13,
    MicrosoftTeamsV1D13Convert.microsoftTeamsV1D13ToJson,
  ],
  "1.14": [
    MicrosoftTeamsV1D14Convert.toMicrosoftTeamsV1D14,
    MicrosoftTeamsV1D14Convert.microsoftTeamsV1D14ToJson,
  ],
  "1.15": [
    MicrosoftTeamsV1D15Convert.toMicrosoftTeamsV1D15,
    MicrosoftTeamsV1D15Convert.microsoftTeamsV1D15ToJson,
  ],
  "1.16": [
    MicrosoftTeamsV1D16Convert.toMicrosoftTeamsV1D16,
    MicrosoftTeamsV1D16Convert.microsoftTeamsV1D16ToJson,
  ],
  "1.17": [
    MicrosoftTeamsV1D17Convert.toMicrosoftTeamsV1D17,
    MicrosoftTeamsV1D17Convert.microsoftTeamsV1D17ToJson,
  ],
  "1.19": [
    MicrosoftTeamsV1D19Convert.toMicrosoftTeamsV1D19,
    MicrosoftTeamsV1D19Convert.microsoftTeamsV1D19ToJson,
  ],
  "1.20": [
    MicrosoftTeamsV1D20Convert.toMicrosoftTeamsV1D20,
    MicrosoftTeamsV1D20Convert.microsoftTeamsV1D20ToJson,
  ],
  devPreview: [
    MicrosoftTeamsVDevPreviewConvert.toMicrosoftTeamsVDevPreview,
    MicrosoftTeamsVDevPreviewConvert.microsoftTeamsVDevPreviewToJson,
  ],
};
const daConverterMap: Converters = {
  "1.0": [
    CopilotDeclarativeAgentV1D0Convert.toCopilotDeclarativeAgentV1D0,
    CopilotDeclarativeAgentV1D0Convert.copilotDeclarativeAgentV1D0ToJson,
  ],
  "1.2": [
    CopilotDeclarativeAgentV1D2Convert.toCopilotDeclarativeAgentV1D2,
    CopilotDeclarativeAgentV1D2Convert.copilotDeclarativeAgentV1D2ToJson,
  ],
  "1.3": [
    CopilotDeclarativeAgentV1D3Convert.toCopilotDeclarativeAgentV1D3,
    CopilotDeclarativeAgentV1D3Convert.copilotDeclarativeAgentV1D3ToJson,
  ],
};
const ApiPluginConverterMap: Converters = {
  "2.1": [
    CopilotPluginV2D1Convert.toCopilotPluginV2D1,
    CopilotPluginV2D1Convert.copilotPluginV2D1ToJson,
  ],
  "2.2": [
    CopilotPluginV2D2Convert.toCopilotPluginV2D2,
    CopilotPluginV2D2Convert.copilotPluginV2D2ToJson,
  ],
};

export class TeamsManifestConverters {
  static jsonToManifest(json: string): MicrosoftTeamsManifest {
    const parsed = JSON.parse(json);
    const manifestVersion = parsed.manifestVersion as string;
    const converters =
      TeamsManifestConverterMap[manifestVersion as keyof typeof TeamsManifestConverterMap];
    if (!converters) {
      throw new Error(
        `Teams manifest version ${manifestVersion} is not supported. Supported versions are: ${Object.keys(
          TeamsManifestConverterMap
        ).join(", ")}`
      );
    }
    return converters[0](json) as MicrosoftTeamsManifest;
  }
  static manifestToJson(manifest: MicrosoftTeamsManifest): string {
    const manifestVersion = manifest.manifestVersion as string;
    const converters =
      TeamsManifestConverterMap[manifestVersion as keyof typeof TeamsManifestConverterMap];
    if (!converters) {
      throw new Error(
        `Teams manifest version ${manifestVersion} is not supported. Supported versions are: ${Object.keys(
          TeamsManifestConverterMap
        ).join(", ")}`
      );
    }
    return converters[1](manifest as any);
  }
}

export class DeclarativeAgentManifestConverters {
  static jsonToManifest(json: string): DeclarativeAgentManifest {
    const parsed = JSON.parse(json);
    const version = parsed.version as string;
    const converters = daConverterMap[version as keyof typeof daConverterMap];
    if (!converters) {
      throw new Error(
        `Declarative agent manifest version ${version} is not supported. Supported versions are: ${Object.keys(
          TeamsManifestConverterMap
        ).join(", ")}`
      );
    }
    return converters[0](json);
  }
  static manifestToJson(manifest: DeclarativeAgentManifest): string {
    const version = manifest.version as string;
    const converters = daConverterMap[version as keyof typeof daConverterMap];
    if (!converters) {
      throw new Error(
        `Declarative agent manifest version ${version} is not supported. Supported versions are: ${Object.keys(
          TeamsManifestConverterMap
        ).join(", ")}`
      );
    }
    return converters[1](manifest);
  }
}

export class ApiPluginManifestConverters {
  static jsonToManifest(json: string): ApiPluginManifest {
    const parsed = JSON.parse(json);
    const schema_version = parsed.schema_version as string;
    const converters =
      ApiPluginConverterMap[schema_version as keyof typeof ApiPluginConverterMap];
    if (!converters) {
      throw new Error(
        `API plugin manifest version ${schema_version} is not supported. Supported versions are: ${Object.keys(
          TeamsManifestConverterMap
        ).join(", ")}`
      );
    }
    return converters[0](json);
  }
  static manifestToJson(manifest: ApiPluginManifest): string {
    const schema_version = manifest.schema_version as string;
    const converters =
      ApiPluginConverterMap[schema_version as keyof typeof ApiPluginConverterMap];
    if (!converters) {
      throw new Error(
        `API plugin manifest version ${schema_version} is not supported. Supported versions are: ${Object.keys(
          TeamsManifestConverterMap
        ).join(", ")}`
      );
    }
    return converters[1](manifest);
  }
}

export class AppManifestUtils {
  static async fetchSchema(manifest: AppManifest): Promise<JSONSchemaType<AppManifest>> {
    const schemaUrl = manifest.$schema as string;
    if (!schemaUrl) {
      throw new Error("Manifest does not have a $schema property or schema url is not provided.");
    }
    let result: JSONSchemaType<AppManifest>;
    try {
      const res = await fetch(schemaUrl);
      result = (await res.json()) as JSONSchemaType<AppManifest>;
    } catch (e: unknown) {
      if (e instanceof Error) {
        throw new Error(`Failed to get manifest at url ${schemaUrl} due to: ${e.message}`);
      } else {
        throw new Error(`Failed to get manifest at url ${schemaUrl} due to: unknown error`);
      }
    }
    return result;
  }
  static async validateAgainstSchema(manifest: AppManifest, schema?: JSONSchemaType<AppManifest>): Promise<string[]> {
    if (!schema) {
      schema = await this.fetchSchema(manifest);
    }
    let validate;
    if (schema.$schema?.includes("2020-12")) {
      const ajv = new Ajv2020({
        //formats: { uri: true, email: true },
        allErrors: true,
        strictTypes: false,
      });
      addFormats(ajv, ["uri", "email", "regex"]);
      validate = ajv.compile(schema);
    } else {
      const ajv = new Ajv({
        allErrors: true,
        strictTypes: false,
      });
      addFormats(ajv, ["uri", "email", "regex"]);
      validate = ajv.compile(schema);
    }
    const valid = validate(manifest);
    if (!valid && validate.errors) {
      return validate.errors.map(
        (error) =>
          `${error.instancePath} ${error.message || ""}. Details: ${
            error.params ? JSON.stringify(error.params) : ""
          }`
      );
    } else {
      return [];
    }
  }

  /**
   * Read manifest from file with basic type check
   *
   * @param filePath - The manifest file path.
   * @throws Will propagate any error thrown by the fs-extra#readFile or type assert failure.
   *
   * @returns The manifest Object
   */
  static async read(filePath: string): Promise<AppManifest> {
    const jsonString = await fs.readFile(filePath, "utf8");
    const manifest = TeamsManifestConverters.jsonToManifest(jsonString);
    return manifest;
  }

  /**
   * Read manifest from file with schema validation
   *
   * @param filePath - The manifest file path.
   * @throws Will propagate any error thrown by the fs-extra#readFile or type check failure.
   *
   * @returns The manifest Object and schema validation results
   */
  static async readAndValidate(filePath: string): Promise<[AppManifest, string[]]> {
    const manifest = await this.read(filePath);
    const validateRes = await this.validateAgainstSchema(manifest);
    return [manifest, validateRes];
  }

  /**
   * Writes the Teams manifest object to the given file.
   *
   * @param path - The manifest file path.
   * @param manifest - Manifest object to be saved
   * @throws Will propagate any error thrown by the fs-extra#writeFile.
   *
   */
  static async writeTeamsManifest(
    filePath: string,
    manifest: MicrosoftTeamsManifest
  ): Promise<void> {
    const jsonString = TeamsManifestConverters.manifestToJson(manifest);
    return fs.writeFile(filePath, jsonString, "utf8");
  }

  /**
   * Writes the declarative agent manifest object to the given file.
   *
   * @param path - The manifest file path.
   * @param manifest - Manifest object to be saved
   * @throws Will propagate any error thrown by the fs-extra#writeFile.
   *
   */
  static async writeDeclarativeAgentManifest(
    filePath: string,
    manifest: DeclarativeAgentManifest
  ): Promise<void> {
    const jsonString = DeclarativeAgentManifestConverters.manifestToJson(manifest);
    return fs.writeFile(filePath, jsonString, "utf8");
  }

  /**
   * Writes the declarative agent manifest object to the given file.
   *
   * @param path - The manifest file path.
   * @param manifest - Manifest object to be saved
   * @throws Will propagate any error thrown by the fs-extra#writeFile.
   *
   */
  static async writeAPIPluginManifest(
    filePath: string,
    manifest: ApiPluginManifest
  ): Promise<void> {
    const jsonString = ApiPluginManifestConverters.manifestToJson(manifest);
    return fs.writeFile(filePath, jsonString, "utf8");
  }
}
