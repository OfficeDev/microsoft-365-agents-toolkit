// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { Result } from "neverthrow";
import type { AtkContext } from "../core/context";
import type { AtkError } from "../core/error";
import type { Question } from "../api/qm/question";
import type { Inputs } from "../api/types";

/**
 * Template categories aligned with the Teams app model.
 */
export type TemplateCategory =
  | "bot"
  | "tab"
  | "message-extension"
  | "custom-engine-agent"
  | "declarative-agent"
  | "ai-agent"
  | "connector"
  | "office-addin";

/**
 * Supported programming languages for templates.
 */
export type Language = "typescript" | "javascript" | "python" | "csharp" | "common";

/**
 * Result of a scaffold operation.
 */
export interface ScaffoldResult {
  projectPath: string;
  warnings?: string[];
}

/**
 * Result of a provision operation.
 */
export interface ProvisionResult {
  outputs: Record<string, string>;
  warnings?: string[];
}

/**
 * Result of a deploy operation.
 */
export interface DeployResult {
  outputs: Record<string, string>;
  warnings?: string[];
}

/**
 * Options passed to scaffold/provision/deploy functions.
 */
export interface TemplateActionOptions {
  /** Target language */
  language: Language;
  /** Project name */
  projectName: string;
  /** Destination folder path */
  destinationPath: string;
  /** Additional template-specific options */
  [key: string]: unknown;
}

/**
 * TemplateDescriptor is a plain data object describing a project template.
 * Replaces the Generator class hierarchy with a flat, registrable descriptor.
 *
 * Templates are plain objects — no classes, no inheritance.
 * `features.json` is the source of truth — descriptors are derived from it.
 */
export interface TemplateDescriptor {
  /** Unique template ID using category/variant convention (e.g. "da/basic", "bot/echo") */
  id: string;

  /** Human-readable display name (e.g. "Echo Bot") */
  name: string;

  /** Short description shown in UI tooltips */
  description?: string;

  /** Template category for grouping in UI */
  category: TemplateCategory;

  /** Supported programming languages */
  languages: Language[];

  /** Scaffold function — creates the project from template */
  scaffoldFn: (
    ctx: AtkContext,
    opts: TemplateActionOptions
  ) => Promise<Result<ScaffoldResult, AtkError>>;

  /** Optional provision function — provisions cloud resources */
  provisionFn?: (
    ctx: AtkContext,
    opts: TemplateActionOptions
  ) => Promise<Result<ProvisionResult, AtkError>>;

  /** Optional deploy function — deploys to cloud */
  deployFn?: (
    ctx: AtkContext,
    opts: TemplateActionOptions
  ) => Promise<Result<DeployResult, AtkError>>;

  /** Additional questions this template needs beyond the common ones */
  questions?: QuestionSpec[];

  /** Sort order within its category (lower appears first) */
  displayOrder?: number;

  /** Feature flag name — template is hidden unless this flag is enabled */
  featureFlag?: string;

  /** Searchable tags for filtering */
  tags?: string[];

  /** ADO test suite ID for traceability */
  adoSuiteId?: number;

  /** Whether this template is testable or tracked-only */
  testable?: boolean;
}

/**
 * QuestionSpec declares an additional input a template needs.
 * Composable — templates list the specs they require, and the question tree builder
 * assembles them into the right position in the IQTreeNode tree.
 */
export interface QuestionSpec {
  /** The question definition (SingleSelect, Text, File, etc.) */
  question: Question;
  /** Name of a parent question this depends on (for conditional activation) */
  dependsOn?: string;
  /** Condition function — when returns true, this question is shown */
  condition?: (inputs: Inputs) => boolean;
}
