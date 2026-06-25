// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import {
  FxError,
  Inputs,
  Platform,
  SystemError,
  UserInteraction,
  err,
} from "@microsoft/teamsfx-api";
import fs from "fs-extra";
import path from "path";
import { Result } from "neverthrow";
import { featureFlagManager } from "../common/featureFlags";
import { TOOLS } from "../common/globalVars";
import { Answers, BuildTarget, DeclarativeLocator, bundledFloorDir, runCreateInputs } from "../v4";
import { runModifySelector } from "../v4/surface/modifySelectorWalk";

const SOURCE = "Scaffold";

export interface ModifyFrontDoorDeps {
  scaffoldV4: (
    inputs: Inputs,
    target: BuildTarget,
    answers: Answers
  ) => Promise<Result<undefined, FxError>>;
  callCoreMethod: (inputs: Inputs, target: BuildTarget) => Promise<Result<undefined, FxError>>;
  flagReader?: (name: string) => boolean;
  readFloorBytes?: () => Buffer;
  ui?: UserInteraction;
  runSelector?: typeof runModifySelector;
  runInputs?: typeof runCreateInputs;
}

function defaultFlagReader(name: string): boolean {
  return featureFlagManager.getBooleanValue({ name, defaultValue: "false" });
}

function readBundledFloorBytes(): Buffer {
  return fs.readFileSync(path.join(bundledFloorDir(), "templates.zip"));
}

function surfaceOf(platform: Platform | undefined): string {
  switch (platform) {
    case Platform.CLI:
    case Platform.CLI_HELP:
      return "cli";
    case Platform.VS:
      return "vs";
    default:
      return "vscode";
  }
}

function unsupportedModifyTarget(target: BuildTarget): SystemError {
  return new SystemError({
    source: SOURCE,
    name: "UnsupportedModifyEngine",
    message: `The modify front door does not dispatch the '${target.engine}' engine.`,
  });
}

function floorReadError(error: unknown): SystemError {
  if (error instanceof SystemError) {
    return error;
  }
  return new SystemError({
    source: SOURCE,
    name: "ModifyTemplatePackageReadFailed",
    message: "Failed to read the bundled modify template package.",
  });
}

export async function modifyProjectFrontDoor(
  inputs: Inputs,
  selectorPrefill: Record<string, string>,
  entryParams: Answers,
  deps: ModifyFrontDoorDeps
): Promise<Result<undefined, FxError>> {
  const flagReader = deps.flagReader ?? defaultFlagReader;
  const ui = deps.ui ?? TOOLS.ui;
  let floorBytes: Buffer;
  try {
    floorBytes = (deps.readFloorBytes ?? readBundledFloorBytes)();
  } catch (error) {
    return err(floorReadError(error));
  }
  const surface = surfaceOf(inputs.platform);
  const runSelector = deps.runSelector ?? runModifySelector;
  const target = await runSelector(floorBytes, ui, surface, {
    flagReader,
    interactive: !inputs.nonInteractive,
    prefilled: selectorPrefill,
  });
  if (target.isErr()) {
    return err(target.error);
  }

  switch (target.value.engine) {
    case "v4": {
      const runInputs = deps.runInputs ?? runCreateInputs;
      const locator: DeclarativeLocator = { kind: "modify", templateId: target.value.templateId };
      const answers = await runInputs(
        floorBytes,
        locator,
        { ...(target.value.answers ?? {}), ...entryParams },
        ui,
        { flagReader, surface }
      );
      if (answers.isErr()) {
        return err(answers.error);
      }
      return deps.scaffoldV4(inputs, target.value, answers.value);
    }
    case "v3-core-method":
      return deps.callCoreMethod(inputs, target.value);
    case "v3":
    case "surface-action":
      return err(unsupportedModifyTarget(target.value));
  }
}

export const modifyProjectFrontDoorDefaults = {
  flagReader: defaultFlagReader,
  readFloorBytes: readBundledFloorBytes,
  runSelector: runModifySelector,
  runInputs: runCreateInputs,
};
