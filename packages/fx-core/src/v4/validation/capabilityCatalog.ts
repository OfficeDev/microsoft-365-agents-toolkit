// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { capabilityDeclarations } from "../capabilities/declarations";
import type { CapabilityDeclaration, CapabilityKind } from "../capabilities/declarations";

function findCapability(kind: CapabilityKind, id: string): CapabilityDeclaration | undefined {
  return Object.values(capabilityDeclarations[kind]).find((declaration) => declaration.id === id);
}

/** Return when a template-visible capability first became available in the engine. */
export function templateCapabilityFloor(
  kind: CapabilityKind,
  id: string,
  output?: string
): string | undefined {
  const capability = findCapability(kind, id);
  return capability?.outputs?.find((entry) => entry.name === output)?.since ?? capability?.since;
}

/** Enumerate a capability kind for registry/catalogue parity tests. */
export function templateCapabilities(kind: CapabilityKind): string[] {
  return Object.values(capabilityDeclarations[kind]).map((declaration) => declaration.id);
}

/** Return the render-context keys a capability may derive. */
export function templateCapabilityOutputs(kind: CapabilityKind, id: string): string[] {
  return findCapability(kind, id)?.outputs?.map((output) => output.name) ?? [];
}
