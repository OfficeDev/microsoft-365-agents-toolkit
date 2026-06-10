// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * OOP Wrappers for Manifest Types
 *
 * This module provides object-oriented wrappers around the generated manifest types,
 * offering a fluent API for manipulation, state tracking, and convenient operations.
 *
 * @module wrappers
 */

export { BaseManifest } from "./BaseManifest";
export { APIPluginManifestWrapper, RuntimeType } from "./APIPluginManifestWrapper";
export type { RuntimeTypeValue } from "./APIPluginManifestWrapper";
export { DeclarativeAgentManifestWrapper, CapabilityName } from "./DeclarativeAgentManifestWrapper";
export type { CapabilityNameValue } from "./DeclarativeAgentManifestWrapper";
export type { AgentSkillElement } from "../generated-types";
export {
  TeamsManifestWrapper,
  DefaultInstallScope,
  StaticTabScope,
  ConfigurableTabScope,
  ComposeExtensionTypeEnum,
} from "./TeamsManifestWrapper";
export type {
  DefaultInstallScopeValue,
  StaticTabScopeValue,
  ConfigurableTabScopeValue,
  ComposeExtensionTypeValue,
} from "./TeamsManifestWrapper";
