// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import Ajv, { JSONSchemaType } from "ajv";
import addFormats from "ajv-formats";
import Ajv2020 from "ajv/dist/2020";
import fs from "fs-extra";
import path from "path";
import {
  DeclarativeAgentManifestV1D0,
  Convert as DeclarativeAgentManifestV1D0Convert,
} from "./copilot/declarative-agent/DeclarativeAgentManifestV1D0";
import {
  DeclarativeAgentManifestV1D2,
  Convert as DeclarativeAgentManifestV1D2Convert,
} from "./copilot/declarative-agent/DeclarativeAgentManifestV1D2";
import {
  DeclarativeAgentManifestV1D3,
  Convert as DeclarativeAgentManifestV1D3Convert,
} from "./copilot/declarative-agent/DeclarativeAgentManifestV1D3";
import {
  APIPluginManifestV2D1,
  Convert as ApiPluginManifestV2D1Convert,
} from "./copilot/plugin/ApiPluginManifestV2D1";
import {
  APIPluginManifestV2D2,
  Convert as ApiPluginManifestV2D2Convert,
} from "./copilot/plugin/ApiPluginManifestV2D2";
import { TeamsManifestV1D0 } from "./teams/TeamsManifestV1D0";
import {
  TeamsManifestV1D1, Convert as TeamsManifestV1D1Convert,
} from "./teams/TeamsManifestV1D1";
import {
  TeamsManifestV1D10,
  Convert as TeamsManifestV1D10Convert,
} from "./teams/TeamsManifestV1D10";
import {
  TeamsManifestV1D11,
  Convert as TeamsManifestV1D11Convert,
} from "./teams/TeamsManifestV1D11";
import {
  TeamsManifestV1D12,
  Convert as TeamsManifestV1D12Convert,
} from "./teams/TeamsManifestV1D12";
import {
  TeamsManifestV1D13,
  Convert as TeamsManifestV1D13Convert,
} from "./teams/TeamsManifestV1D13";
import {
  TeamsManifestV1D14,
  Convert as TeamsManifestV1D14Convert,
} from "./teams/TeamsManifestV1D14";
import {
  TeamsManifestV1D15,
  Convert as TeamsManifestV1D15Convert,
} from "./teams/TeamsManifestV1D15";
import {
  TeamsManifestV1D16,
  Convert as TeamsManifestV1D16Convert,
} from "./teams/TeamsManifestV1D16";
import {
  TeamsManifestV1D17,
  Convert as TeamsManifestV1D17Convert,
} from "./teams/TeamsManifestV1D17";
import {
  TeamsManifestV1D19,
  Convert as TeamsManifestV1D19Convert,
} from "./teams/TeamsManifestV1D19";
import {
  TeamsManifestV1D2,
  Convert as TeamsManifestV1D2Convert,
} from "./teams/TeamsManifestV1D2";
import {
  TeamsManifestV1D20,
  Convert as TeamsManifestV1D20Convert,
} from "./teams/TeamsManifestV1D20";
import {
  TeamsManifestV1D3,
  Convert as TeamsManifestV1D3Convert,
} from "./teams/TeamsManifestV1D3";
import {
  TeamsManifestV1D4,
  Convert as TeamsManifestV1D4Convert,
} from "./teams/TeamsManifestV1D4";
import {
  TeamsManifestV1D5,
  Convert as TeamsManifestV1D5Convert,
} from "./teams/TeamsManifestV1D5";
import {
  TeamsManifestV1D6,
  Convert as TeamsManifestV1D6Convert,
} from "./teams/TeamsManifestV1D6";
import {
  TeamsManifestV1D7,
  Convert as TeamsManifestV1D7Convert,
} from "./teams/TeamsManifestV1D7";
import {
  TeamsManifestV1D8,
  Convert as TeamsManifestV1D8Convert,
} from "./teams/TeamsManifestV1D8";
import {
  TeamsManifestV1D9,
  Convert as TeamsManifestV1D9Convert,
} from "./teams/TeamsManifestV1D9";
import {
  TeamsManifestVDevPreview,
  Convert as TeamsManifestVDevPreviewConvert,
} from "./teams/TeamsManifestVDevPreview";

export {
  TeamsManifestV1D1,
  Convert as TeamsManifestV1D1Convert
} from "./teams/TeamsManifestV1D1";
export {
  TeamsManifestV1D10,
  Convert as TeamsManifestV1D10Convert
} from "./teams/TeamsManifestV1D10";
export {
  TeamsManifestV1D11,
  Convert as TeamsManifestV1D11Convert
} from "./teams/TeamsManifestV1D11";
export {
  TeamsManifestV1D12,
  Convert as TeamsManifestV1D12Convert
} from "./teams/TeamsManifestV1D12";
export {
  TeamsManifestV1D13,
  Convert as TeamsManifestV1D13Convert
} from "./teams/TeamsManifestV1D13";
export {
  TeamsManifestV1D14,
  Convert as TeamsManifestV1D14Convert
} from "./teams/TeamsManifestV1D14";
export {
  TeamsManifestV1D15,
  Convert as TeamsManifestV1D15Convert
} from "./teams/TeamsManifestV1D15";
export {
  TeamsManifestV1D16,
  Convert as TeamsManifestV1D16Convert
} from "./teams/TeamsManifestV1D16";
export {
  TeamsManifestV1D17,
  Convert as TeamsManifestV1D17Convert
} from "./teams/TeamsManifestV1D17";
export {
  TeamsManifestV1D19,
  Convert as TeamsManifestV1D19Convert
} from "./teams/TeamsManifestV1D19";
export {
  TeamsManifestV1D2,
  Convert as TeamsManifestV1D2Convert
} from "./teams/TeamsManifestV1D2";
export {
  TeamsManifestV1D20,
  Convert as TeamsManifestV1D20Convert
} from "./teams/TeamsManifestV1D20";
export {
  TeamsManifestV1D3,
  Convert as TeamsManifestV1D3Convert
} from "./teams/TeamsManifestV1D3";
export {
  TeamsManifestV1D4,
  Convert as TeamsManifestV1D4Convert
} from "./teams/TeamsManifestV1D4";
export {
  TeamsManifestV1D5,
  Convert as TeamsManifestV1D5Convert
} from "./teams/TeamsManifestV1D5";
export {
  TeamsManifestV1D6,
  Convert as TeamsManifestV1D6Convert
} from "./teams/TeamsManifestV1D6";
export {
  TeamsManifestV1D7,
  Convert as TeamsManifestV1D7Convert
} from "./teams/TeamsManifestV1D7";
export {
  TeamsManifestV1D8,
  Convert as TeamsManifestV1D8Convert
} from "./teams/TeamsManifestV1D8";
export {
  TeamsManifestV1D9,
  Convert as TeamsManifestV1D9Convert
} from "./teams/TeamsManifestV1D9";
export { TeamsManifestVDevPreview as DevPreviewSchema, TeamsManifestVDevPreview, Convert as TeamsManifestVDevPreviewConvert } from "./teams/TeamsManifestVDevPreview";
export type TeamsManifest =
  | TeamsManifestV1D0 & { manifestVersion: "1.0"; $schema?: string }
  | TeamsManifestV1D1 & { manifestVersion: "1.1" }
  | TeamsManifestV1D2 & { manifestVersion: "1.2" }
  | TeamsManifestV1D3 & { manifestVersion: "1.3" }
  | TeamsManifestV1D4 & { manifestVersion: "1.4" }
  | TeamsManifestV1D5
  | TeamsManifestV1D6
  | TeamsManifestV1D7
  | TeamsManifestV1D8
  | TeamsManifestV1D9
  | TeamsManifestV1D10
  | TeamsManifestV1D11
  | TeamsManifestV1D12
  | TeamsManifestV1D13
  | TeamsManifestV1D14
  | TeamsManifestV1D15
  | TeamsManifestV1D16
  | TeamsManifestV1D17
  | TeamsManifestV1D19
  | TeamsManifestV1D20
  | TeamsManifestVDevPreview;

export type TeamsManifestLatest = TeamsManifestV1D20;

export {
  DeclarativeAgentManifestV1D0,
  Convert as DeclarativeAgentManifestV1D0Convert
} from "./copilot/declarative-agent/DeclarativeAgentManifestV1D0";
export {
  DeclarativeAgentManifestV1D2,
  Convert as DeclarativeAgentManifestV1D2Convert
} from "./copilot/declarative-agent/DeclarativeAgentManifestV1D2";
export {
  DeclarativeAgentManifestV1D3,
  Convert as DeclarativeAgentManifestV1D3Convert
} from "./copilot/declarative-agent/DeclarativeAgentManifestV1D3";
export type DeclarativeAgentManifest =
  | DeclarativeAgentManifestV1D0
  | DeclarativeAgentManifestV1D2
  | DeclarativeAgentManifestV1D3;

export type DeclarativeAgentManifestLatest = DeclarativeAgentManifestV1D3;

export {
  APIPluginManifestV2D1,
  Convert as ApiPluginV2D1Convert
} from "./copilot/plugin/ApiPluginManifestV2D1";
export {
  APIPluginManifestV2D2,
  Convert as ApiPluginManifestV2D2Convert
} from "./copilot/plugin/ApiPluginManifestV2D2";
export type APIPluginManifest = APIPluginManifestV2D1 | APIPluginManifestV2D2;
export type APIPluginManifestLatest = APIPluginManifestV2D2;

export type AppManifest = TeamsManifest | DeclarativeAgentManifest | APIPluginManifest;

type Converters = {
  [key: string]: [(json: string) => any, (manifest: any) => string];
};
const TeamsManifestConverterMap: Converters = {
  "1.1": [
    TeamsManifestV1D1Convert.toTeamsManifestV1D1,
    TeamsManifestV1D1Convert.teamsManifestV1D1ToJson,
  ],
  "1.2": [
    TeamsManifestV1D2Convert.toTeamsManifestV1D2,
    TeamsManifestV1D2Convert.teamsManifestV1D2ToJson,
  ],
  "1.3": [
    TeamsManifestV1D3Convert.toTeamsManifestV1D3,
    TeamsManifestV1D3Convert.teamsManifestV1D3ToJson,
  ],
  "1.4": [
    TeamsManifestV1D4Convert.toTeamsManifestV1D4,
    TeamsManifestV1D4Convert.teamsManifestV1D4ToJson,
  ],
  "1.5": [
    TeamsManifestV1D5Convert.toTeamsManifestV1D5,
    TeamsManifestV1D5Convert.teamsManifestV1D5ToJson,
  ],
  "1.6": [
    TeamsManifestV1D6Convert.toTeamsManifestV1D6,
    TeamsManifestV1D6Convert.teamsManifestV1D6ToJson,
  ],
  "1.7": [
    TeamsManifestV1D7Convert.toTeamsManifestV1D7,
    TeamsManifestV1D7Convert.teamsManifestV1D7ToJson,
  ],
  "1.8": [
    TeamsManifestV1D8Convert.toTeamsManifestV1D8,
    TeamsManifestV1D8Convert.teamsManifestV1D8ToJson,
  ],
  "1.9": [
    TeamsManifestV1D9Convert.toTeamsManifestV1D9,
    TeamsManifestV1D9Convert.teamsManifestV1D9ToJson,
  ],
  "1.10": [
    TeamsManifestV1D10Convert.toTeamsManifestV1D10,
    TeamsManifestV1D10Convert.teamsManifestV1D10ToJson,
  ],
  "1.11": [
    TeamsManifestV1D11Convert.toTeamsManifestV1D11,
    TeamsManifestV1D11Convert.teamsManifestV1D11ToJson,
  ],
  "1.12": [
    TeamsManifestV1D12Convert.toTeamsManifestV1D12,
    TeamsManifestV1D12Convert.teamsManifestV1D12ToJson,
  ],
  "1.13": [
    TeamsManifestV1D13Convert.toTeamsManifestV1D13,
    TeamsManifestV1D13Convert.teamsManifestV1D13ToJson,
  ],
  "1.14": [
    TeamsManifestV1D14Convert.toTeamsManifestV1D14,
    TeamsManifestV1D14Convert.teamsManifestV1D14ToJson,
  ],
  "1.15": [
    TeamsManifestV1D15Convert.toTeamsManifestV1D15,
    TeamsManifestV1D15Convert.teamsManifestV1D15ToJson,
  ],
  "1.16": [
    TeamsManifestV1D16Convert.toTeamsManifestV1D16,
    TeamsManifestV1D16Convert.teamsManifestV1D16ToJson,
  ],
  "1.17": [
    TeamsManifestV1D17Convert.toTeamsManifestV1D17,
    TeamsManifestV1D17Convert.teamsManifestV1D17ToJson,
  ],
  "1.19": [
    TeamsManifestV1D19Convert.toTeamsManifestV1D19,
    TeamsManifestV1D19Convert.teamsManifestV1D19ToJson,
  ],
  "1.20": [
    TeamsManifestV1D20Convert.toTeamsManifestV1D20,
    TeamsManifestV1D20Convert.teamsManifestV1D20ToJson,
  ],
  devPreview: [
    TeamsManifestVDevPreviewConvert.toTeamsManifestVDevPreview,
    TeamsManifestVDevPreviewConvert.teamsManifestVDevPreviewToJson,
  ],
};
const daConverterMap: Converters = {
  "1.0": [
    DeclarativeAgentManifestV1D0Convert.toDeclarativeAgentManifestV1D0,
    DeclarativeAgentManifestV1D0Convert.declarativeAgentManifestV1D0ToJson,
  ],
  "1.2": [
    DeclarativeAgentManifestV1D2Convert.toDeclarativeAgentManifestV1D2,
    DeclarativeAgentManifestV1D2Convert.declarativeAgentManifestV1D2ToJson,
  ],
  "1.3": [
    DeclarativeAgentManifestV1D3Convert.toDeclarativeAgentManifestV1D3,
    DeclarativeAgentManifestV1D3Convert.declarativeAgentManifestV1D3ToJson,
  ],
};
const ApiPluginConverterMap: Converters = {
  "2.1": [
    ApiPluginManifestV2D1Convert.toAPIPluginManifestV2D1,
    ApiPluginManifestV2D1Convert.aPIPluginManifestV2D1ToJson,
  ],
  "2.2": [
    ApiPluginManifestV2D2Convert.toAPIPluginManifestV2D2,
    ApiPluginManifestV2D2Convert.aPIPluginManifestV2D2ToJson,
  ],
};

export class TeamsManifestConverter {
  static jsonToManifest(json: string): TeamsManifest {
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
    return converters[0](json) as TeamsManifest;
  }
  static manifestToJson(manifest: TeamsManifest): string {
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

export class DeclarativeAgentManifestConverter {
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

export class ApiPluginManifestConverter {
  static jsonToManifest(json: string): APIPluginManifest {
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
  static manifestToJson(manifest: APIPluginManifest): string {
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
  /**
   * Fecth the schema from the manifest object, load from local if the schema is in the package
   * @param manifest 
   * @returns manifest schema object
   */
  static async fetchSchema(schemaUrl: string): Promise<JSONSchemaType<AppManifest>> {
    if (schemaUrl.startsWith("https://developer.microsoft.com/json-schemas/teams") 
      || schemaUrl.startsWith("https://developer.microsoft.com/json-schemas/copilot/declarative-agent") 
      || schemaUrl.startsWith("https://developer.microsoft.com/json-schemas/copilot/plugin")) {
      const suffix = schemaUrl.substring("https://developer.microsoft.com/".length);
      const schemaFile = path.join(__dirname, "..", suffix);
      if (await fs.pathExists(schemaFile)) {
        const json = await fs.readJson(schemaFile);
        return json as JSONSchemaType<AppManifest>;
      }
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
      const schemaUrl = manifest.$schema;
      if (!schemaUrl) {
        throw new Error("Manifest does not have a $schema property");
      }
      schema = await this.fetchSchema(schemaUrl);
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
   * Read Teams manifest from file with basic type check
   *
   * @param filePath - Teams manifest file path.
   * @throws Will propagate any error thrown by the fs-extra#readFile or type assert failure.
   *
   * @returns The manifest Object
   */
  static async readTeamsManifest(filePath: string): Promise<TeamsManifest> {
    const jsonString = await fs.readFile(filePath, "utf8");
    const manifest = TeamsManifestConverter.jsonToManifest(jsonString);
    return manifest;
  }

  /**
   * Read declarative agent manifest from file with schema validation
   *
   * @param filePath - Teams manifest file path.
   * @throws Will propagate any error thrown by the fs-extra#readFile or type check failure.
   *
   * @returns The manifest Object and schema validation results
   */
  static async readAndValidateTeamsManifest(filePath: string): Promise<[TeamsManifest, string[]]> {
    const manifest = await this.readTeamsManifest(filePath);
    const validateRes = await this.validateAgainstSchema(manifest);
    return [manifest, validateRes];
  }

  /**
   * Read declarative agent manifest from file with basic type check
   *
   * @param filePath - Declarative agent manifest file path.
   * @throws Will propagate any error thrown by the fs-extra#readFile or type assert failure.
   *
   * @returns The manifest Object
   */
  static async readDeclarativeAgentManifest(filePath: string): Promise<DeclarativeAgentManifest> {
    const jsonString = await fs.readFile(filePath, "utf8");
    const manifest = DeclarativeAgentManifestConverter.jsonToManifest(jsonString);
    return manifest;
  }

  /**
   * Read declarative agent manifest from file with schema validation
   *
   * @param filePath - Declarative agent manifest file path.
   * @throws Will propagate any error thrown by the fs-extra#readFile or type check failure.
   *
   * @returns The manifest Object and schema validation results
   */
  static async readAndValidateDeclarativeAgentManifest(filePath: string): Promise<[DeclarativeAgentManifest, string[]]> {
    const manifest = await this.readDeclarativeAgentManifest(filePath);
    const validateRes = await this.validateAgainstSchema(manifest);
    return [manifest, validateRes];
  }

  /**
   * Read API plugin manifest from file with basic type check
   *
   * @param filePath - API plugin manifest file path.
   * @throws Will propagate any error thrown by the fs-extra#readFile or type assert failure.
   *
   * @returns The manifest Object
   */
  static async readApiPluginManifest(filePath: string): Promise<APIPluginManifest> {
    const jsonString = await fs.readFile(filePath, "utf8");
    const manifest = ApiPluginManifestConverter.jsonToManifest(jsonString);
    return manifest;
  }
  
  /**
   * Read API plugin manifest from file with schema validation
   *
   * @param filePath - API plugin manifest file path.
   * @throws Will propagate any error thrown by the fs-extra#readFile or type check failure.
   *
   * @returns The manifest Object and schema validation results
   */
  static async readAndValidateApiPluginManifest(filePath: string): Promise<[APIPluginManifest, string[]]> {
    const manifest = await this.readApiPluginManifest(filePath);
    const validateRes = await this.validateAgainstSchema(manifest);
    return [manifest, validateRes];
  }

  /**
   * Writes the Teams manifest object to the given file with basic type check.
   *
   * @param path - The manifest file path.
   * @param manifest - Manifest object to be saved
   * @throws Will propagate any error thrown by the fs-extra#writeFile.
   *
   */
  static async writeTeamsManifest(
    filePath: string,
    manifest: TeamsManifest
  ): Promise<void> {
    const jsonString = TeamsManifestConverter.manifestToJson(manifest);
    return fs.writeFile(filePath, jsonString, "utf8");
  }

  /**
   * Writes the declarative agent manifest object to the given file with basic type check.
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
    const jsonString = DeclarativeAgentManifestConverter.manifestToJson(manifest);
    return fs.writeFile(filePath, jsonString, "utf8");
  }

  /**
   * Writes the declarative agent manifest object to the given file with basic type check.
   *
   * @param path - The manifest file path.
   * @param manifest - Manifest object to be saved
   * @throws Will propagate any error thrown by the fs-extra#writeFile.
   *
   */
  static async writeApiPluginManifest(
    filePath: string,
    manifest: APIPluginManifest
  ): Promise<void> {
    const jsonString = ApiPluginManifestConverter.manifestToJson(manifest);
    return fs.writeFile(filePath, jsonString, "utf8");
  }
}
