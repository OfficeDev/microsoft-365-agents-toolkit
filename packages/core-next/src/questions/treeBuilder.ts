// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import type { IQTreeNode } from "../api/qm/question";
import type { OptionItem, Inputs } from "../api/types";
import type { TemplateRegistry } from "../templates/registry";
import type { TemplateCategory, TemplateDescriptor, QuestionSpec } from "../templates/types";
import type { FeatureFlagRegistry } from "../featureFlags";
import {
  projectNameQuestion,
  destinationFolderQuestion,
  languageQuestion,
  projectTypeQuestion,
  templateIdQuestion,
} from "./commonQuestions";

/**
 * Category display names and order.
 */
const categoryMeta: Record<TemplateCategory, { label: string; order: number }> = {
  "declarative-agent": { label: "Declarative Agent", order: 1 },
  "custom-engine-agent": { label: "Custom Engine Agent", order: 2 },
  "ai-agent": { label: "AI Agent", order: 3 },
  bot: { label: "Bot", order: 4 },
  tab: { label: "Tab", order: 5 },
  "message-extension": { label: "Message Extension", order: 6 },
  connector: { label: "Connector", order: 7 },
  "office-addin": { label: "Office Add-in", order: 8 },
};

/**
 * Build a full question tree from the TemplateRegistry.
 *
 * Tree structure:
 *   root (group)
 *     ├── projectType (singleSelect) — categories from registry
 *     │   └── [per category] templateId (singleSelect) — templates in that category
 *     │       └── [per template] language (singleSelect) — languages for that template
 *     │           └── [per template] ...extra QuestionSpec[] from descriptor
 *     ├── projectName (text)
 *     └── destinationFolder (folder)
 */
export function buildQuestionTree(
  registry: TemplateRegistry,
  featureFlags?: FeatureFlagRegistry
): IQTreeNode {
  const allDescriptors = registry
    .list()
    .filter((d) => !d.featureFlag || featureFlags?.isEnabled(d.featureFlag));

  // Group by category
  const grouped = groupByCategory(allDescriptors);

  // Build category options
  const categoryOptions: OptionItem[] = [];
  const categoryChildren: IQTreeNode[] = [];

  for (const [cat, descriptors] of grouped) {
    const meta = categoryMeta[cat] ?? { label: cat, order: 99 };
    categoryOptions.push({
      id: cat,
      label: meta.label,
    });

    // Build template selector for this category
    const templateOptions: OptionItem[] = descriptors
      .sort((a, b) => (a.displayOrder ?? 99) - (b.displayOrder ?? 99))
      .map((d) => ({
        id: d.id,
        label: d.name,
        description: d.description,
      }));

    const templateChildren: IQTreeNode[] = descriptors
      .sort((a, b) => (a.displayOrder ?? 99) - (b.displayOrder ?? 99))
      .map((d) => buildTemplateNode(d));

    const templateNode: IQTreeNode = {
      data: templateIdQuestion(templateOptions),
      condition: { equals: cat },
      children: templateChildren,
    };

    categoryChildren.push(templateNode);
  }

  // Sort categories
  categoryOptions.sort(
    (a, b) =>
      (categoryMeta[a.id as TemplateCategory]?.order ?? 99) -
      (categoryMeta[b.id as TemplateCategory]?.order ?? 99)
  );

  const root: IQTreeNode = {
    data: { type: "group", name: "createProject" },
    children: [
      {
        data: projectTypeQuestion(categoryOptions),
        children: categoryChildren,
      },
      { data: projectNameQuestion() },
      { data: destinationFolderQuestion() },
    ],
  };

  return root;
}

/**
 * Build the subtree for a single template descriptor.
 * Includes the language selector + any template-specific QuestionSpecs.
 */
function buildTemplateNode(descriptor: TemplateDescriptor): IQTreeNode {
  const children: IQTreeNode[] = [];

  // Language selector (if more than "common")
  if (descriptor.languages.length > 0) {
    children.push({
      data: languageQuestion(descriptor.languages),
    });
  }

  // Template-specific questions
  if (descriptor.questions) {
    for (const spec of descriptor.questions) {
      children.push(questionSpecToNode(spec));
    }
  }

  return {
    data: { type: "group", name: descriptor.id },
    condition: { equals: descriptor.id },
    children: children.length > 0 ? children : undefined,
  };
}

/**
 * Convert a QuestionSpec into an IQTreeNode.
 */
function questionSpecToNode(spec: QuestionSpec): IQTreeNode {
  const node: IQTreeNode = {
    data: spec.question,
  };
  if (spec.condition) {
    const conditionFn = spec.condition;
    node.condition = (inputs: Inputs) => conditionFn(inputs);
  }
  return node;
}

/**
 * Group descriptors by category, preserving insertion order.
 */
function groupByCategory(
  descriptors: TemplateDescriptor[]
): Map<TemplateCategory, TemplateDescriptor[]> {
  const grouped = new Map<TemplateCategory, TemplateDescriptor[]>();
  for (const d of descriptors) {
    const list = grouped.get(d.category) ?? [];
    list.push(d);
    grouped.set(d.category, list);
  }
  // Sort map entries by category order
  const sorted = new Map(
    [...grouped.entries()].sort(
      ([a], [b]) => (categoryMeta[a]?.order ?? 99) - (categoryMeta[b]?.order ?? 99)
    )
  );
  return sorted;
}
