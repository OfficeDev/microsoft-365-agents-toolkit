// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { FxError, SystemError, UserError, UserInteraction } from "@microsoft/teamsfx-api";
import { Result, err, ok } from "neverthrow";
import { SelectorPresentation } from "../buildTarget/parseSelector";
import {
  BUILD_TARGET_UNKNOWN_TEMPLATE,
  BuildTarget,
  RouteResolverPort,
  SelectorWalkResult,
  resolveBuildTarget,
  v4RouteRegistryFromSelector,
} from "../buildTarget/resolveBuildTarget";
import { WalkHistoryEntry } from "../collectInputs/collectInputs";
import {
  openCreateSelector,
  openCreateSelectorPresentation,
  openSelectorFromJsonBytes,
  openSelectorPresentationFromJsonBytes,
} from "../distribution/createSelector";
import { openDeclarativePackage } from "../distribution/declarativePackage";
import { readBooleanFeatureFlag } from "../../common/featureFlags";
import { createSelectorPrompt } from "./selectorPresentation";

/** Live Q1 create-selector prompt face. See walk-create-selector spec. */

const SOURCE = "Scaffold";

/** Create-selector options; all are defaulted. */
export interface CreateSelectorDeps {
  /** The feature-flag reader (default: env-backed); v4 imports no `featureFlagManager`. */
  flagReader?: (name: string) => boolean;
  /** Selector bytes shape. Defaults to the full package zip for current callers. */
  selectorBytesKind?: "zip" | "json";
  /** Membership test supplied by a staged artifact snapshot or metadata index. */
  v4Registry?: (templateId: string) => boolean;
  /** Q1 answers known up front. */
  prefilled?: Record<string, string>;
  /** Whether unfilled required dimensions may be prompted. */
  interactive?: boolean;
  /** Resume a prior Q1 walk (cross-phase back): re-ask its last dimension with the retained history. */
  resume?: { history: WalkHistoryEntry[] };
}

/** The default env-backed feature-flag reader (a flag is on iff its env var is exactly `"true"`). */
function envFlagReader(name: string): boolean {
  return readBooleanFeatureFlag(name);
}

/** Convert a thrown prompt failure back to an `FxError` for the `Result` boundary. */
function toFxError(e: unknown): FxError {
  if (e instanceof UserError || e instanceof SystemError) {
    return e;
  }
  const message = e instanceof Error ? e.message : String(e);
  return new SystemError({ source: SOURCE, name: "CreateSelectorWalkFailed", message });
}

/** Build the live interactive route resolver over the floor and host UI. */
function buildPort(
  floorBytes: Buffer,
  presentation: SelectorPresentation,
  ui: UserInteraction,
  surface: string,
  flagReader: (name: string) => boolean,
  v4Registry: ((templateId: string) => boolean) | undefined
): RouteResolverPort {
  return {
    prompt: createSelectorPrompt(presentation, ui, surface, flagReader),
    featureFlag: flagReader,
    v4Registry(templateId: string): boolean {
      if (v4Registry !== undefined) {
        return v4Registry(templateId);
      }
      return openDeclarativePackage(floorBytes, { kind: "create", templateId }).isOk();
    },
  };
}

/** Run create Q1 over `ui`, resolving the dispatched `BuildTarget`. */
export async function runCreateSelector(
  floorBytes: Buffer,
  ui: UserInteraction,
  surface: string,
  deps: CreateSelectorDeps = {}
): Promise<Result<SelectorWalkResult, FxError>> {
  const flagReader = deps.flagReader ?? envFlagReader;
  const selectorBytesKind = deps.selectorBytesKind ?? "zip";
  const prefilled = deps.prefilled ?? {};
  const interactive = deps.interactive ?? true;
  const spec =
    selectorBytesKind === "json"
      ? openSelectorFromJsonBytes(floorBytes, "create")
      : openCreateSelector(floorBytes);
  if (spec.isErr()) {
    return err(spec.error);
  }
  const presentation =
    selectorBytesKind === "json"
      ? openSelectorPresentationFromJsonBytes(floorBytes, "create")
      : openCreateSelectorPresentation(floorBytes);
  if (presentation.isErr()) {
    return err(presentation.error);
  }
  const port = buildPort(
    floorBytes,
    presentation.value,
    ui,
    surface,
    flagReader,
    deps.v4Registry ??
      (selectorBytesKind === "json" ? v4RouteRegistryFromSelector(spec.value) : undefined)
  );
  try {
    return await resolveBuildTarget(spec.value, prefilled, interactive, port, {
      resume: deps.resume,
    });
  } catch (e) {
    return err(toFxError(e));
  }
}

/** Resolve a pinned template id without re-walking Q1. */
export function resolveCreateTargetByTemplateId(
  floorBytes: Buffer,
  templateId: string
): Result<BuildTarget, FxError> {
  const spec = openCreateSelector(floorBytes);
  if (spec.isErr()) {
    return err(spec.error);
  }
  const route = spec.value.routes.find((r) => r.templateId === templateId);
  if (route === undefined) {
    return err(
      new UserError({
        source: SOURCE,
        name: BUILD_TARGET_UNKNOWN_TEMPLATE,
        message: `Template '${templateId}' is not present in the create selector.`,
      })
    );
  }
  return ok({ templateId, engine: route.engine, answers: {} });
}
