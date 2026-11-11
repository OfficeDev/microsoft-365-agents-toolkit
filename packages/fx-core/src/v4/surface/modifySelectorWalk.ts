// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { FxError, SystemError, UserError, UserInteraction } from "@microsoft/teamsfx-api";
import { Result, err } from "neverthrow";
import { SelectorPresentation } from "../buildTarget/parseSelector";
import {
  BuildTarget,
  RouteResolverPort,
  resolveBuildTarget,
  v4RouteRegistryFromSelector,
} from "../buildTarget/resolveBuildTarget";
import {
  openModifySelector,
  openModifySelectorPresentation,
  openSelectorFromJsonBytes,
  openSelectorPresentationFromJsonBytes,
} from "../distribution/createSelector";
import { openDeclarativePackage } from "../distribution/declarativePackage";
import { readBooleanFeatureFlag } from "../../common/featureFlags";
import { createSelectorPrompt } from "./selectorPresentation";

const SOURCE = "Scaffold";

export interface ModifySelectorDeps {
  flagReader?: (name: string) => boolean;
  selectorBytesKind?: "zip" | "json";
  v4Registry?: (templateId: string) => boolean;
  prefilled?: Record<string, string>;
  interactive?: boolean;
}

function envFlagReader(name: string): boolean {
  return readBooleanFeatureFlag(name);
}

function toFxError(e: unknown): FxError {
  if (e instanceof UserError || e instanceof SystemError) {
    return e;
  }
  const message = e instanceof Error ? e.message : String(e);
  return new SystemError({ source: SOURCE, name: "ModifySelectorWalkFailed", message });
}

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
      return openDeclarativePackage(floorBytes, { kind: "modify", templateId }).isOk();
    },
  };
}

export async function runModifySelector(
  floorBytes: Buffer,
  ui: UserInteraction,
  surface: string,
  deps: ModifySelectorDeps = {}
): Promise<Result<BuildTarget, FxError>> {
  const flagReader = deps.flagReader ?? envFlagReader;
  const selectorBytesKind = deps.selectorBytesKind ?? "zip";
  const prefilled = deps.prefilled ?? {};
  const interactive = deps.interactive ?? true;
  const spec =
    selectorBytesKind === "json"
      ? openSelectorFromJsonBytes(floorBytes, "modify")
      : openModifySelector(floorBytes);
  if (spec.isErr()) {
    return err(spec.error);
  }
  const presentation =
    selectorBytesKind === "json"
      ? openSelectorPresentationFromJsonBytes(floorBytes, "modify")
      : openModifySelectorPresentation(floorBytes);
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
    return await resolveBuildTarget(spec.value, prefilled, interactive, port);
  } catch (e) {
    return err(toFxError(e));
  }
}
