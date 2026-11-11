// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import {
  DeclarativeAgentManifestWrapper,
  FxError,
  SystemError,
  TeamsManifestWrapper,
} from "@microsoft/teamsfx-api";
import { Result, err, ok } from "neverthrow";
import path from "path";
import type { StepContext } from "../../pipeline/runScaffoldPipeline";

export interface DaManifestService {
  registerDeclarativeAgentAction(
    io: Pick<StepContext, "read" | "write">,
    teamsManifestPath: string,
    pluginManifestPath: string
  ): Result<void, FxError>;
  setSensitivityLabel(
    io: Pick<StepContext, "read" | "write">,
    manifestPath: string,
    id: string
  ): Result<void, FxError>;
}

function manifestError(name: string, message: string, cause?: unknown): SystemError {
  return cause instanceof Error
    ? new SystemError({ source: "Scaffold", name, message, error: cause })
    : new SystemError({ source: "Scaffold", name, message });
}

function normalizeTargetPath(entryPath: string): string {
  return entryPath.replace(/\\/g, "/");
}

function resolveSiblingPath(baseFile: string, relativeFile: string): string {
  return normalizeTargetPath(
    path.posix.normalize(path.posix.join(path.posix.dirname(baseFile), relativeFile))
  );
}

function pluginFileRelativeToAgent(agentManifestPath: string, pluginManifestPath: string): string {
  return normalizeTargetPath(
    path.posix.relative(path.posix.dirname(agentManifestPath), pluginManifestPath)
  );
}

function actionId(pluginManifestPath: string): string {
  const basename = path.posix.basename(pluginManifestPath, path.posix.extname(pluginManifestPath));
  return basename.startsWith("ai-plugin-") ? basename.substring("ai-plugin-".length) : basename;
}

function readManifest(
  io: Pick<StepContext, "read" | "write">,
  filePath: string,
  missingErrorName: string,
  readErrorName: string
): Result<Buffer, FxError> {
  try {
    const contents = io.read(filePath);
    return contents === undefined
      ? err(manifestError(missingErrorName, `Cannot read '${filePath}'.`))
      : ok(contents);
  } catch (error) {
    return err(manifestError(readErrorName, `Cannot read '${filePath}'.`, error));
  }
}

function registerDeclarativeAgentAction(
  io: Pick<StepContext, "read" | "write">,
  teamsManifestPath: string,
  pluginManifestPath: string
): Result<void, FxError> {
  const teamsManifestContents = readManifest(
    io,
    teamsManifestPath,
    "DaActionTeamsManifestMissing",
    "DaActionTeamsManifestReadFailed"
  );
  if (teamsManifestContents.isErr()) {
    return err(teamsManifestContents.error);
  }

  let teamsManifest: TeamsManifestWrapper;
  try {
    teamsManifest = TeamsManifestWrapper.fromJSON(teamsManifestContents.value.toString("utf8"));
  } catch {
    return err(
      manifestError("DaActionTeamsManifestInvalid", `'${teamsManifestPath}' is not valid JSON.`)
    );
  }
  const agentFile = teamsManifest.getDeclarativeAgentPaths()[0];
  if (agentFile === undefined) {
    return err(
      manifestError(
        "DaActionManifestFileMissing",
        `The Teams manifest '${teamsManifestPath}' does not reference a declarative agent manifest.`
      )
    );
  }

  const agentManifestPath = resolveSiblingPath(teamsManifestPath, agentFile);
  const agentManifestContents = readManifest(
    io,
    agentManifestPath,
    "DaActionManifestMissing",
    "DaActionManifestReadFailed"
  );
  if (agentManifestContents.isErr()) {
    return err(agentManifestContents.error);
  }

  let agentManifest: DeclarativeAgentManifestWrapper;
  try {
    agentManifest = DeclarativeAgentManifestWrapper.fromJSON(
      agentManifestContents.value.toString("utf8")
    );
  } catch {
    return err(
      manifestError("DaActionManifestInvalid", `'${agentManifestPath}' is not valid JSON.`)
    );
  }
  agentManifest.upsertAction(
    actionId(pluginManifestPath),
    pluginFileRelativeToAgent(agentManifestPath, pluginManifestPath)
  );
  try {
    io.write(agentManifestPath, Buffer.from(agentManifest.toJSON(), "utf8"));
  } catch (error) {
    return err(
      manifestError("DaActionManifestWriteFailed", `Cannot write '${agentManifestPath}'.`, error)
    );
  }
  return ok(undefined);
}

function setSensitivityLabel(
  io: Pick<StepContext, "read" | "write">,
  manifestPath: string,
  id: string
): Result<void, FxError> {
  const contents = readManifest(
    io,
    manifestPath,
    "DaSensitivityLabelManifestMissing",
    "DaSensitivityLabelManifestReadFailed"
  );
  if (contents.isErr()) {
    return err(contents.error);
  }

  let serialized: Buffer;
  try {
    const wrapper = DeclarativeAgentManifestWrapper.fromJSON(contents.value.toString("utf8"));
    wrapper.setSensitivityLabel(id);
    serialized = Buffer.from(wrapper.toJSON(), "utf8");
  } catch (error) {
    return err(
      manifestError(
        "DaSensitivityLabelManifestInvalid",
        "the Declarative Agent manifest to label is invalid",
        error
      )
    );
  }
  try {
    io.write(manifestPath, serialized);
    return ok(undefined);
  } catch (error) {
    return err(
      manifestError(
        "DaSensitivityLabelManifestWriteFailed",
        `Cannot write '${manifestPath}'.`,
        error
      )
    );
  }
}

export const daManifestService: DaManifestService = {
  registerDeclarativeAgentAction,
  setSensitivityLabel,
};
