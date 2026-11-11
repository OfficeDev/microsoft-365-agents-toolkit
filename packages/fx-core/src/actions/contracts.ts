// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { FxError, Result, UserError, err, ok } from "@microsoft/teamsfx-api";
import { getLocalizedString } from "../common/localizeUtils";
import { acquireActionName, acquireArgs } from "../component/driver/m365/args";
import { publishPolicyActionName, publishPolicyArgs } from "../component/driver/workiq/args";
import {
  ActionArgsValidationError,
  ArgsContract,
  validationIssue,
} from "../component/driver/util/args";

export {
  ActionArgsValidationError,
  ActionArgsValidationIssue,
} from "../component/driver/util/args";

export const ActionNames = Object.freeze({
  Acquire: acquireActionName,
  PublishPolicy: publishPolicyActionName,
});

const catalogue = {
  [ActionNames.Acquire]: { ...acquireArgs, supportsCancellation: false },
  [ActionNames.PublishPolicy]: { ...publishPolicyArgs, supportsCancellation: true },
};

export type ActionName = keyof typeof catalogue;
type ArgsOf<T> = T extends ArgsContract<infer TArgs> ? TArgs : never;
export type ActionArgs<K extends ActionName> = ArgsOf<(typeof catalogue)[K]>;

const contracts: {
  [K in ActionName]: ArgsContract<ActionArgs<K>> & { supportsCancellation: boolean };
} = catalogue;

export interface ActionDefinition {
  name: ActionName;
  argsSchema: Record<string, unknown>;
  supportsCancellation: boolean;
}

const actionErrorKeys = {
  UnsupportedAction: "core.actions.UnsupportedAction",
  InvalidActionContext: "core.actions.InvalidActionContext",
  UnsupportedCancellation: "core.actions.UnsupportedCancellation",
};

export function createActionError(name: keyof typeof actionErrorKeys): FxError {
  return new UserError({
    source: "executeAction",
    name,
    message: getLocalizedString(actionErrorKeys[name]),
    displayMessage: getLocalizedString(actionErrorKeys[name]),
  });
}

function isActionName(name: unknown): name is ActionName {
  return typeof name === "string" && Object.hasOwn(catalogue, name);
}

export function getActionDefinition(name: string): Result<ActionDefinition, FxError> {
  if (!isActionName(name)) return err(createActionError("UnsupportedAction"));
  const contract = contracts[name];
  return ok({
    name,
    argsSchema: structuredClone(contract.argsSchema),
    supportsCancellation: contract.supportsCancellation,
  });
}

export function validateActionArgs<K extends ActionName>(
  name: K,
  input: unknown
): Result<ActionArgs<K>, ActionArgsValidationError> {
  if (!isActionName(name)) {
    return err(new ActionArgsValidationError([validationIssue("", "unknownAction")]));
  }
  return contracts[name].validate(input);
}
