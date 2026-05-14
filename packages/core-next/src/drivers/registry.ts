// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { DriverDescriptor } from "./types";

/**
 * DriverRegistry is a lookup-based registry for all available drivers.
 * Drivers self-register at module load time via `register()`.
 *
 * Replaces the old DriverContext/container with a flat, discoverable registry.
 */
export class DriverRegistry {
  private readonly drivers = new Map<string, DriverDescriptor>();

  /**
   * Register a driver descriptor. Throws if ID is already registered.
   */
  register(descriptor: DriverDescriptor): void {
    if (this.drivers.has(descriptor.id)) {
      throw new Error(
        `Driver "${descriptor.id}" is already registered. Duplicate driver IDs are not allowed.`
      );
    }
    this.drivers.set(descriptor.id, descriptor);
  }

  /**
   * Get a driver by ID. Returns undefined if not found.
   */
  get(id: string): DriverDescriptor | undefined {
    return this.drivers.get(id);
  }

  /**
   * Check if a driver ID is registered.
   */
  has(id: string): boolean {
    return this.drivers.has(id);
  }

  /**
   * Get all registered driver descriptors.
   */
  list(): DriverDescriptor[] {
    return Array.from(this.drivers.values());
  }

  /**
   * Get the count of registered drivers.
   */
  get size(): number {
    return this.drivers.size;
  }
}

/**
 * Global driver registry instance.
 * Each driver module calls `driverRegistry.register(descriptor)` at load.
 */
export const driverRegistry = new DriverRegistry();
