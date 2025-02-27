// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author yefuwang@microsoft.com
 */

import { FxError, Result, ok, err } from "@microsoft/teamsfx-api";
import fs from "fs-extra";
import { load } from "js-yaml";
import { globalVars } from "../../common/globalVars";
import { InvalidYamlSchemaError, YamlFieldMissingError, YamlFieldTypeError } from "../../error/yml";
import {
  IYamlParser,
  ProjectModel,
  RawProjectModel,
  LifecycleNames,
  AdditionalMetadata,
} from "./interface";
import { Lifecycle } from "./lifecycle";
import { Validator } from "./validator";
import { getLocalizedString } from "../../common/localizeUtils";

const validator = new Validator();

const environmentFolderPath = "environmentFolderPath";
const writeToEnvironmentFile = "writeToEnvironmentFile";
const versionNotSupportedKey = "error.yaml.VersionNotSupported";

function parseRawProjectModel(obj: Record<string, any>): Result<RawProjectModel, FxError> {
  const result: RawProjectModel = { version: "" };
  if (environmentFolderPath in obj) {
    if (typeof obj[environmentFolderPath] !== "string") {
      return err(new YamlFieldTypeError("environmentFolderPath", "string"));
    }
    result.environmentFolderPath = obj[environmentFolderPath];
  }

  if ("version" in obj) {
    if (typeof obj["version"] !== "string") {
      return err(new YamlFieldTypeError("version", "string"));
    }
    result.version = obj["version"];
  } else {
    return err(new YamlFieldMissingError("version"));
  }

  if ("additionalMetadata" in obj) {
    // No validation for additionalMetadata by design. This property is for telemetry related purpose only
    // and should not affect user-observable behavior of TTK.
    result.additionalMetadata = obj["additionalMetadata"] as AdditionalMetadata;
  }

  for (const name of LifecycleNames) {
    if (name in obj) {
      const value = obj[name];
      if (!Array.isArray(value)) {
        return err(new YamlFieldTypeError(name, "array"));
      }
      for (const elem of value) {
        if (!("uses" in elem)) {
          return err(new YamlFieldMissingError(`${name}.uses`));
        }
        if (!(typeof elem["uses"] === "string")) {
          return err(new YamlFieldTypeError(`${name}.uses`, "string"));
        }
        if (!("with" in elem)) {
          return err(new YamlFieldMissingError(`${name}.with`));
        }
        if (!(typeof elem["with"] === "object")) {
          return err(new YamlFieldTypeError(`${name}.with`, "object"));
        }
        if (elem["env"]) {
          if (typeof elem["env"] !== "object" || Array.isArray(elem["env"])) {
            return err(new YamlFieldTypeError(`${name}.env`, "object"));
          }
          for (const envVar in elem["env"]) {
            if (typeof elem["env"][envVar] !== "string") {
              return err(new YamlFieldTypeError(`${name}.env.${envVar}`, "string"));
            }
          }
        }
        if (elem[writeToEnvironmentFile]) {
          if (
            typeof elem[writeToEnvironmentFile] !== "object" ||
            Array.isArray(elem[writeToEnvironmentFile])
          ) {
            return err(new YamlFieldTypeError(`${name}.writeToEnvironmentFile`, "object"));
          }
          for (const envVar in elem[writeToEnvironmentFile]) {
            if (typeof elem[writeToEnvironmentFile][envVar] !== "string") {
              return err(
                new YamlFieldTypeError(`${name}.writeToEnvironmentFile.${envVar}`, "string")
              );
            }
          }
        }
      }
      result[name] = value;
    }
  }

  return ok(result);
}

export class YamlParser implements IYamlParser {
  async parse(path: string, validateSchema?: boolean): Promise<Result<ProjectModel, FxError>> {
    const raw = await this.parseRaw(path, validateSchema);
    if (raw.isErr()) {
      return err(raw.error);
    }
    const result: ProjectModel = { version: raw.value.version };
    for (const name of LifecycleNames) {
      if (name in raw.value) {
        const definitions = raw.value[name];
        if (definitions) {
          result[name] = new Lifecycle(name, definitions, result.version);
        }
      }
    }

    if (raw.value.environmentFolderPath) {
      result.environmentFolderPath = raw.value.environmentFolderPath;
    }

    if (raw.value.additionalMetadata) {
      result.additionalMetadata = raw.value.additionalMetadata;
    }

    return ok(result);
  }

  private async parseRaw(
    path: string,
    validateSchema?: boolean
  ): Promise<Result<RawProjectModel, FxError>> {
    try {
      globalVars.ymlFilePath = path;
      const str = await fs.readFile(path, "utf8");
      const content = load(str);
      const value = content as unknown as Record<string, unknown>;
      const version = typeof value["version"] === "string" ? value["version"] : undefined;
      // note: typeof null === "object" typeof undefined === "undefined" in js
      if (typeof content !== "object" || Array.isArray(content) || content === null) {
        return err(new InvalidYamlSchemaError(path));
      }
      if (validateSchema) {
        if (!validator.isVersionSupported(version ?? "undefined")) {
          return err(
            new InvalidYamlSchemaError(
              path,
              getLocalizedString(
                versionNotSupportedKey,
                version,
                validator.supportedVersions().join(", ")
              )
            )
          );
        }
        const valid = validator.validate(value, version);
        if (!valid) {
          return err(new InvalidYamlSchemaError(path));
        }
      }
      return parseRawProjectModel(value);
    } catch (error) {
      return err(new InvalidYamlSchemaError(path));
    }
  }
}

export const yamlParser = new YamlParser();
