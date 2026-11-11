// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { Result, UserError, err, ok } from "@microsoft/teamsfx-api";
import { Ajv, ErrorObject, JSONSchemaType } from "ajv";
import { getLocalizedString } from "../../../common/localizeUtils";

const ajv = new Ajv({ allErrors: true, ownProperties: true });

export interface ArgsContract<T> {
  argsSchema: Record<string, unknown>;
  validate(input: unknown): Result<T, ActionArgsValidationError>;
}

export interface ActionArgsValidationIssue {
  path: string;
  code: string;
  message: string;
}

export class ActionArgsValidationError extends UserError {
  public constructor(public readonly issues: ActionArgsValidationIssue[]) {
    super({
      source: "executeAction",
      name: "ActionArgsValidationError",
      message: getLocalizedString("core.actions.invalidArgs"),
      displayMessage: getLocalizedString("core.actions.invalidArgs"),
    });
  }
}

export function validationIssue(path: string, code: string): ActionArgsValidationIssue {
  return { path, code, message: getLocalizedString("core.actions.invalidField") };
}

function toIssue(error: ErrorObject): ActionArgsValidationIssue {
  const property: unknown =
    error.keyword === "required"
      ? error.params.missingProperty
      : error.keyword === "additionalProperties"
        ? error.params.additionalProperty
        : undefined;
  const path =
    typeof property === "string"
      ? `${error.instancePath}/${property.replace(/~/g, "~0").replace(/\//g, "~1")}`
      : error.instancePath;
  return validationIssue(path, error.keyword);
}

export function createArgsContract<T>(argsSchema: JSONSchemaType<T>): ArgsContract<T> {
  const schema: Record<string, unknown> = {
    $schema: "http://json-schema.org/draft-07/schema#",
  };
  Object.assign(schema, argsSchema);
  const validator = ajv.compile<T>(schema);
  return {
    argsSchema: schema,
    validate(input) {
      let snapshot: unknown;
      try {
        snapshot = structuredClone(input);
      } catch {
        return err(new ActionArgsValidationError([validationIssue("", "type")]));
      }
      if (!validator(snapshot)) {
        return err(new ActionArgsValidationError((validator.errors ?? []).map(toIssue)));
      }
      return ok(snapshot);
    },
  };
}
