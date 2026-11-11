// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { FxError, Warning } from "@microsoft/teamsfx-api";
import { Result } from "neverthrow";
import { RenderVars } from "../model/dataModel";
import { ExpressionRuntimePort, Scope, evaluateExpression } from "../expression/evaluateExpression";
import {
  Orchestration,
  PipelineRuntimePort,
  RegisteredStep,
} from "../pipeline/runScaffoldPipeline";
import { renderMustache } from "./renderMustache";
import {
  STEP_REGISTER_PLUGIN_MANIFEST,
  createDaActionRegisterPluginManifestStep,
} from "./steps/daAction";
import { DaManifestService, daManifestService } from "./services/daManifestService";
import {
  GeneralSensitivityLabelService,
  NOOP_GENERAL_SENSITIVITY_LABEL_SERVICE,
  STEP_SET_SENSITIVITY_LABEL,
  createDaSetSensitivityLabelStep,
} from "./steps/daSensitivity";
import {
  STEP_INJECT_YML_ACTION,
  STEP_PERSIST_CREDENTIAL_ENV,
  mcpAuthInjectYmlAction,
  mcpAuthPersistCredentialEnv,
} from "./steps/mcpAuth";
import { STEP_MATERIALIZE_LOCAL_SERVERS, mcpLocalMaterializeServers } from "./steps/mcpLocal";
import { STEP_MATERIALIZE_STATIC_MCP_TOOLS, mcpStaticMaterializeTools } from "./steps/mcpStatic";
import {
  STEP_IMPORT_EXISTING_OFFICE_ADDIN_PROJECT,
  officeAddinImportExistingProject,
} from "./steps/officeAddin";
import {
  STEP_GENERATE_OPENAPI_PLUGIN_FILES,
  STEP_GENERATE_TEAMS_AI_CUSTOM_API_FILES,
  openApiGeneratePluginFiles,
  openApiGenerateTeamsAiCustomApiFiles,
} from "./steps/openApi";
import {
  STEP_UNIFY_PROJECT_ID,
  STEP_UPGRADE_EXISTING_PROJECT,
  metaOsUnifyProjectId,
  metaOsUpgradeExistingProject,
} from "./steps/metaOs";
/** Shared v4 pipeline registry and port factory. See ADR-0017 for whitelist rules. */

/** The orchestration names the engine knows (ADR-0017 closed whitelist). */
export const KNOWN_PIPELINES = new Set(["default", "openapi", "typespec", "officeAddin", "spfx"]);

/** Generic named-step lookup assembled at a runtime composition boundary. */
export type StepRegistry = ReadonlyMap<string, RegisteredStep>;

/** Bind runtime-owned business adapters into the closed post-render step whitelist. */
export function createStepRegistry(
  generalSensitivityLabel: GeneralSensitivityLabelService = NOOP_GENERAL_SENSITIVITY_LABEL_SERVICE,
  manifests: DaManifestService = daManifestService
): StepRegistry {
  return new Map<string, RegisteredStep>([
    [STEP_REGISTER_PLUGIN_MANIFEST, createDaActionRegisterPluginManifestStep(manifests)],
    [
      STEP_SET_SENSITIVITY_LABEL,
      createDaSetSensitivityLabelStep(generalSensitivityLabel, manifests),
    ],
    [STEP_INJECT_YML_ACTION, mcpAuthInjectYmlAction],
    [STEP_PERSIST_CREDENTIAL_ENV, mcpAuthPersistCredentialEnv],
    [STEP_MATERIALIZE_LOCAL_SERVERS, mcpLocalMaterializeServers],
    [STEP_MATERIALIZE_STATIC_MCP_TOOLS, mcpStaticMaterializeTools],
    [STEP_GENERATE_OPENAPI_PLUGIN_FILES, openApiGeneratePluginFiles],
    [STEP_GENERATE_TEAMS_AI_CUSTOM_API_FILES, openApiGenerateTeamsAiCustomApiFiles],
    [STEP_IMPORT_EXISTING_OFFICE_ADDIN_PROJECT, officeAddinImportExistingProject],
    [STEP_UNIFY_PROJECT_ID, metaOsUnifyProjectId],
    [STEP_UPGRADE_EXISTING_PROJECT, metaOsUpgradeExistingProject],
  ]);
}

/** Default post-render step whitelist for offline runtimes and compatibility tests. */
export const STEP_REGISTRY = createStepRegistry();

/** Runtime-specific file sink injected behind the shared pipeline port. */
export interface FileSink {
  writeNew(path: string, data: Buffer): boolean;
  /** Persist `data` at `path` (a target-relative, forward-slash path). */
  write(path: string, data: Buffer): void;
  /** Read back a previously written file, or `undefined` when absent (EAFP). */
  read(path: string): Buffer | undefined;
}

/** Runtime-owned persistence boundary for regular and secret environment values. */
export type EnvironmentWriter = (
  environment: string,
  values: Record<string, string>
) => Promise<Result<void, FxError>>;

/** Drop list-valued render vars before calling the scalar expression evaluator. */
function scalarScope(renderVars: RenderVars): Scope {
  const scope: Scope = {};
  for (const [key, value] of Object.entries(renderVars)) {
    if (Array.isArray(value)) {
      continue;
    }
    scope[key] = value;
  }
  return scope;
}

/** Build the shared pipeline port over an injected file sink. */
export function buildPipelinePort(
  exprPort: ExpressionRuntimePort,
  sink: FileSink,
  environmentWriter: EnvironmentWriter,
  stepRegistry: StepRegistry = STEP_REGISTRY,
  warningSink?: (warning: Warning) => void
): PipelineRuntimePort {
  return {
    pipelineRegistry: (name: string): Orchestration | undefined =>
      KNOWN_PIPELINES.has(name) ? { name } : undefined,
    stepRegistry: (name: string): RegisteredStep | undefined => stepRegistry.get(name),
    evalWhen: (expr: string, renderVars: RenderVars): Result<boolean, FxError> =>
      evaluateExpression({ expr }, scalarScope(renderVars), exprPort).map(
        (value) => value === true
      ),
    render: (mustache: string, renderVars: RenderVars): Result<string, FxError> =>
      renderMustache(mustache, renderVars),
    warn: warningSink,
    writeNew: (path: string, data: Buffer): boolean => sink.writeNew(path, data),
    write: (path: string, data: Buffer): void => sink.write(path, data),
    writeEnvironment: environmentWriter,
    read: (path: string): Buffer | undefined => sink.read(path),
  };
}
