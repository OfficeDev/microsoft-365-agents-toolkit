// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { JSONSchemaType } from "ajv";
import Ajv from "ajv-draft-04";
import addFormats from "ajv-formats";
import Ajv2020 from "ajv/dist/2020";
import fs from "fs-extra";
import fetch from "node-fetch";
import { ManifestCommonProperties } from "./ManifestCommonProperties";
import { SharePointAppId } from "./constants";
import { DeclarativeCopilotManifestSchema } from "./declarativeCopilotManifest";
import {
  AppManifestUtils,
  DevPreviewSchema,
  MicrosoftTeamsManifest,
  TeamsManifestConverters,
} from "./generated-types";
import { IComposeExtension, TeamsAppManifest } from "./manifest";
import { PluginManifestSchema } from "./pluginManifest";

export * from "./declarativeCopilotManifest";
export * from "./generated-types";
export * from "./manifest";
export * from "./pluginManifest";

export type TeamsAppManifestJSONSchema = JSONSchemaType<TeamsAppManifest>;
export type DevPreviewManifestJSONSchema = JSONSchemaType<DevPreviewSchema>;

/**
 * @deprecated
 */
export type Manifest = TeamsAppManifest | DevPreviewSchema;

export type ManifestProperties = ManifestCommonProperties;

export class ManifestUtil {
  /**
   * Loads the manifest from the given path with basic type check.
   *
   * @deprecated use `AppManifestUtils.read()` instead
   * @param path - The path to the manifest file.
   * @throws Will propagate any error thrown by the fs-extra#readJson or type check failure.
   *
   * @returns The Manifest Object.
   */
  static async loadFromPath(path: string): Promise<MicrosoftTeamsManifest> {
    const jsonString = await fs.readFile(path, "utf8");
    const manifest = TeamsManifestConverters.jsonToManifest(jsonString);
    return manifest;
  }

  /**
   * Loads the manifest from the given path with validation
   *
   * @deprecated use `AppManifestUtils.readAndValidate()` instead
   * @param path - The path to the manifest file.
   * @throws Will propagate any error thrown by the fs-extra#readJson or type check failure.
   *
   * @returns The Manifest Object and schema validation results
   */
  static async loadAndValidateFromPath(path: string): Promise<[MicrosoftTeamsManifest, string[]]> {
    const manifest = await this.loadFromPath(path);
    const validateRes = await AppManifestUtils.validateAgainstSchema(manifest);
    return [manifest, validateRes];
  }

  /**
   * Writes the manifest object to the given path.
   *
   * @deprecated use `AppManifestUtils.writeTeamsManifest()` instead
   * @param path - Where to write
   * @param manifest - Manifest object to be saved
   * @throws Will propagate any error thrown by the fs-extra#writeJson.
   *
   */
  static async writeToPath<T extends Manifest = TeamsAppManifest>(
    path: string,
    manifest: T
  ): Promise<void> {
    return fs.writeJson(path, manifest, { spaces: 4 });
  }

  /**
   * Validate manifest against json schema.
   * @deprecated use `AppManifestUtils.validateAgainstSchema(manifest, schema)` instead
   * @param manifest - Manifest object to be validated
   * @param schema - teams-app-manifest schema
   * @returns An empty array if it passes validation, or an array of error string otherwise.
   */
  static validateManifestAgainstSchema<
    T extends
      | Manifest
      | DeclarativeCopilotManifestSchema
      | PluginManifestSchema
      | MicrosoftTeamsManifest = TeamsAppManifest
  >(manifest: T, schema: JSONSchemaType<T>): Promise<string[]> {
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
      return Promise.resolve(
        validate.errors?.map(
          (error) =>
            `${error.instancePath} ${error.message || ""}. Details: ${
              error.params ? JSON.stringify(error.params) : ""
            }`
        )
      );
    } else {
      return Promise.resolve([]);
    }
  }

  /**
   * @deprecated
   * @param manifest
   * @returns
   */
  static async fetchSchema<
    T extends
      | Manifest
      | DeclarativeCopilotManifestSchema
      | PluginManifestSchema
      | MicrosoftTeamsManifest = TeamsAppManifest
  >(manifest: T): Promise<JSONSchemaType<T>> {
    const schemaUrl = ((manifest as any).$schema || (manifest as any).schema) as string;
    if (!schemaUrl) {
      throw new Error("Manifest does not have a $schema property or schema url is not provided.");
    }
    let result: JSONSchemaType<T>;
    try {
      const res = await fetch(schemaUrl);
      result = (await res.json()) as JSONSchemaType<T>;
    } catch (e: unknown) {
      if (e instanceof Error) {
        throw new Error(`Failed to get manifest at url ${schemaUrl} due to: ${e.message}`);
      } else {
        throw new Error(`Failed to get manifest at url ${schemaUrl} due to: unknown error`);
      }
    }
    return result;
  }

  /**
   * Validate manifest against {@link TeamsAppManifest#$schema}.
   *
   * @deprecated use `AppManifestUtils.validateAgainstSchema(manifest: T)` instead
   * @param manifest - Manifest object to be validated
   * @throws Will throw if {@link TeamsAppManifest#$schema} is undefined, not valid
   *         or there is any network failure when getting the schema.
   *
   * @returns An empty array if schema validation passes, or an array of error string otherwise.
   */
  static async validateManifest<
    T extends
      | Manifest
      | DeclarativeCopilotManifestSchema
      | PluginManifestSchema
      | MicrosoftTeamsManifest = TeamsAppManifest
  >(manifest: T): Promise<string[]> {
    const schema = await this.fetchSchema(manifest);
    return ManifestUtil.validateManifestAgainstSchema(manifest, schema);
  }

  /**
   * Parse the manifest and get properties
   * @param manifest
   */
  static parseCommonProperties<T extends Manifest | MicrosoftTeamsManifest = TeamsAppManifest>(
    manifest: T
  ): ManifestCommonProperties {
    const capabilities: string[] = [];
    if (manifest.staticTabs && manifest.staticTabs.length > 0) {
      capabilities.push("staticTab");
    }
    if (manifest.configurableTabs && manifest.configurableTabs.length > 0) {
      capabilities.push("configurableTab");
    }
    if (manifest.bots && manifest.bots.length > 0) {
      capabilities.push("Bot");
    }
    if (manifest.composeExtensions && manifest.composeExtensions.length > 0) {
      capabilities.push("MessageExtension");
    }

    const properties: ManifestCommonProperties = {
      id: manifest.id,
      version: manifest.version,
      capabilities: capabilities,
      manifestVersion: manifest.manifestVersion,
      isApiME: false,
      isSPFx: false,
      isApiMeAAD: false,
    };

    // If it's copilot plugin app
    if (
      manifest.composeExtensions &&
      manifest.composeExtensions.length > 0 &&
      (manifest.composeExtensions[0] as IComposeExtension).composeExtensionType == "apiBased"
    ) {
      properties.isApiME = true;
    }

    // If it's SPFx app
    if (
      (manifest as any).webApplicationInfo &&
      (manifest as any).webApplicationInfo.id &&
      (manifest as any).webApplicationInfo.id == SharePointAppId
    ) {
      properties.isSPFx = true;
    }

    // If it's API ME with AAD auth
    if (
      manifest.composeExtensions &&
      manifest.composeExtensions.length > 0 &&
      (manifest.composeExtensions[0] as IComposeExtension).composeExtensionType == "apiBased" &&
      (manifest.composeExtensions[0] as IComposeExtension).authorization?.authType ==
        "microsoftEntra"
    ) {
      properties.isApiMeAAD = true;
    }

    if ((manifest as TeamsAppManifest).copilotExtensions?.plugins) {
      const apiPlugins = (manifest as TeamsAppManifest).copilotExtensions?.plugins;
      if (apiPlugins && apiPlugins.length > 0 && apiPlugins[0].file) capabilities.push("plugin");
    }

    if ((manifest as TeamsAppManifest).copilotExtensions?.declarativeCopilots) {
      const copilotGpts = (manifest as TeamsAppManifest).copilotExtensions?.declarativeCopilots;
      if (copilotGpts && copilotGpts.length > 0) capabilities.push("copilotGpt");
    }

    if ((manifest as TeamsAppManifest).copilotAgents?.plugins) {
      const apiPlugins = (manifest as TeamsAppManifest).copilotAgents?.plugins;
      if (
        apiPlugins &&
        apiPlugins.length > 0 &&
        apiPlugins[0].file &&
        !capabilities.includes("plugin")
      )
        capabilities.push("plugin");
    }

    if ((manifest as TeamsAppManifest).copilotAgents?.declarativeAgents) {
      const copilotGpts = (manifest as TeamsAppManifest).copilotAgents?.declarativeAgents;
      if (copilotGpts && copilotGpts.length > 0 && !capabilities.includes("copilotGpt"))
        capabilities.push("copilotGpt");
    }

    return properties;
  }

  /**
   * Parse the manifest and get telemetry propreties e.g. appId, capabilities etc.
   * @param manifest
   * @returns Telemetry properties
   */
  static parseCommonTelemetryProperties(manifest: TeamsAppManifest | MicrosoftTeamsManifest): {
    [p: string]: string;
  } {
    const properties = ManifestUtil.parseCommonProperties(manifest);

    const telemetryProperties: { [p: string]: string } = {};
    const propertiesMap = new Map<string, any>(Object.entries(properties));
    propertiesMap.forEach((value, key) => {
      if (Array.isArray(value)) {
        telemetryProperties[key] = value.join(";");
      } else {
        telemetryProperties[key] = value;
      }
    });

    return telemetryProperties;
  }

  static async useCopilotExtensionsInSchema(
    manifest: TeamsAppManifest | MicrosoftTeamsManifest
  ): Promise<boolean> {
    const schema = await this.fetchSchema(manifest);
    return !!schema.properties.copilotExtensions;
  }
}
