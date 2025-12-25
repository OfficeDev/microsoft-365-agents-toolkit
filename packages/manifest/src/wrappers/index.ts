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
export { PluginManifestWrapper, RuntimeType } from "./PluginManifestWrapper";
export type { RuntimeTypeValue } from "./PluginManifestWrapper";
export { DeclarativeAgentManifestWrapper, CapabilityName } from "./DeclarativeAgentManifestWrapper";
export type { CapabilityNameValue } from "./DeclarativeAgentManifestWrapper";
