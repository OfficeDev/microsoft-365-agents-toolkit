// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import fs from "fs-extra";

/**
 * Abstract base class for manifest wrappers.
 * Provides common functionality for reading, writing, and tracking changes to manifest files.
 *
 * @template T - The manifest type this wrapper handles
 */
export abstract class BaseManifest<T extends Record<string, unknown>> {
  protected _data: T;
  protected _filePath?: string;
  protected _isDirty = false;

  protected constructor(data: T, filePath?: string) {
    this._data = data;
    this._filePath = filePath;
  }

  /**
   * Returns the raw manifest data.
   */
  get data(): Readonly<T> {
    return this._data;
  }

  /**
   * Returns the file path if the manifest was loaded from a file.
   */
  get filePath(): string | undefined {
    return this._filePath;
  }

  /**
   * Indicates whether the manifest has unsaved changes.
   */
  get isDirty(): boolean {
    return this._isDirty;
  }

  /**
   * Marks the manifest as having unsaved changes.
   */
  protected markDirty(): void {
    this._isDirty = true;
  }

  /**
   * Saves the manifest to the specified file path or the original file path.
   * @param filePath - Optional path to save to. If not provided, uses the original file path.
   * @throws Error if no file path is available.
   */
  async save(filePath?: string): Promise<void> {
    const targetPath = filePath ?? this._filePath;
    if (!targetPath) {
      throw new Error("No file path specified for saving.");
    }
    await fs.writeFile(targetPath, this.toJSON(), "utf-8");
    this._filePath = targetPath;
    this._isDirty = false;
  }

  /**
   * Validates the manifest against its schema.
   * @returns Array of validation error messages, empty if valid.
   */
  abstract validate(): Promise<string[]>;

  /**
   * Converts the manifest to a JSON string.
   */
  abstract toJSON(): string;

  /**
   * Creates a deep clone of this manifest wrapper.
   */
  abstract clone(): BaseManifest<T>;

  // ============= Static Helpers =============

  /**
   * Reads a JSON file and returns the parsed object.
   * @param filePath - Path to the JSON file
   */
  protected static async readJsonFile<U>(filePath: string): Promise<U> {
    const content = await fs.readFile(filePath, "utf-8");
    return JSON.parse(content) as U;
  }

  /**
   * Reads a JSON file synchronously and returns the parsed object.
   * @param filePath - Path to the JSON file
   */
  protected static readJsonFileSync<U>(filePath: string): U {
    const content = fs.readFileSync(filePath, "utf-8");
    return JSON.parse(content) as U;
  }
}
