// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { FeatureFlag, FeatureFlagSource, envSource } from "./types";

/**
 * Injectable feature flag registry.
 * Register flags, query their state — testable with any FeatureFlagSource.
 */
export class FeatureFlagRegistry {
  private readonly flags = new Map<string, FeatureFlag>();
  private readonly source: FeatureFlagSource;

  constructor(source?: FeatureFlagSource) {
    this.source = source ?? envSource;
  }

  /** Register a feature flag definition. */
  register(flag: FeatureFlag): void {
    this.flags.set(flag.name, flag);
  }

  /** Register multiple flags at once. */
  registerAll(flags: FeatureFlag[]): void {
    for (const f of flags) {
      this.flags.set(f.name, f);
    }
  }

  /**
   * Check whether a flag is enabled.
   * Reads the env var via the configured source.
   * Returns the default if the env var is unset.
   */
  isEnabled(name: string): boolean {
    const flag = this.flags.get(name);
    if (!flag) return false;
    const raw = this.source.get(name);
    if (raw === undefined || raw === "") return flag.defaultValue;
    return raw.toLowerCase() === "true" || raw === "1";
  }

  /** Get the raw string value of a flag's env var, or undefined. */
  getValue(name: string): string | undefined {
    return this.source.get(name);
  }

  /** List all registered flags. */
  list(): FeatureFlag[] {
    return Array.from(this.flags.values());
  }

  /** List only the flags that are currently enabled. */
  listEnabled(): FeatureFlag[] {
    return this.list().filter((f) => this.isEnabled(f.name));
  }
}
