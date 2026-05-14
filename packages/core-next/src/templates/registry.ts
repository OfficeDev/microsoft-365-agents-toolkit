// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { TemplateDescriptor, Language, TemplateCategory } from "./types";

/**
 * TemplateRegistry is a lookup-based registry for all available templates.
 * Templates self-register at module load time via `register()`.
 *
 * Replaces the old GeneratorProvider with a flat, discoverable registry.
 */
export class TemplateRegistry {
  private readonly templates = new Map<string, TemplateDescriptor>();

  /**
   * Register a template descriptor. Throws if ID is already registered.
   */
  register(descriptor: TemplateDescriptor): void {
    if (this.templates.has(descriptor.id)) {
      throw new Error(
        `Template "${descriptor.id}" is already registered. Duplicate template IDs are not allowed.`
      );
    }
    this.templates.set(descriptor.id, descriptor);
  }

  /**
   * Get a template by ID. Returns undefined if not found.
   */
  get(id: string): TemplateDescriptor | undefined {
    return this.templates.get(id);
  }

  /**
   * Check if a template ID is registered.
   */
  has(id: string): boolean {
    return this.templates.has(id);
  }

  /**
   * Get all registered template descriptors.
   */
  list(): TemplateDescriptor[] {
    return Array.from(this.templates.values());
  }

  /**
   * Filter templates by category.
   */
  listByCategory(category: TemplateCategory): TemplateDescriptor[] {
    return this.list().filter((t) => t.category === category);
  }

  /**
   * Filter templates by supported language.
   */
  listByLanguage(language: Language): TemplateDescriptor[] {
    return this.list().filter((t) => t.languages.includes(language));
  }

  /**
   * Filter templates that have a specific tag.
   */
  listByTag(tag: string): TemplateDescriptor[] {
    return this.list().filter((t) => t.tags?.includes(tag));
  }

  /**
   * Get the count of registered templates.
   */
  get size(): number {
    return this.templates.size;
  }
}

/**
 * Global template registry instance.
 * Each template module calls `templateRegistry.register(descriptor)` at load.
 */
export const templateRegistry = new TemplateRegistry();
