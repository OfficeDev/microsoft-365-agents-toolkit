// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * A feature flag definition.
 */
export interface FeatureFlag {
  /** Environment variable name (e.g. "TEAMSFX_V4_CORE") */
  name: string;
  /** Default when the env var is not set */
  defaultValue: boolean;
  /** Human-readable description */
  description?: string;
}

/**
 * Abstracts how feature flag values are read.
 * Defaults to process.env, but can be swapped for testing.
 */
export interface FeatureFlagSource {
  get(name: string): string | undefined;
}

/**
 * Default source that reads from process.env.
 */
export const envSource: FeatureFlagSource = {
  get(name: string): string | undefined {
    return process.env[name];
  },
};
