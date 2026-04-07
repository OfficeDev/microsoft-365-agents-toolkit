// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import {
  Inputs,
  IQTreeNode,
  OptionItem,
  Platform,
  Question,
  SingleSelectQuestion,
} from "@microsoft/teamsfx-api";
import { featureFlagManager } from "../../common/featureFlags";
import { getLocalizedString } from "../../common/localizeUtils";
import {
  apiSpecNode,
  apiSpecWithSearchNode,
  azureOpenAINode,
  foundryNode,
  llmServiceNode,
} from "./commonNodes";
import { setTemplateName, setTemplateNameAndGC } from "./vsc/CapabilityOptions";
import { GCConnectionIdQuestion, GCNameQuestion } from "../create";
import { QuestionNames } from "../questionNames";

function isFeatureEnabled(flagName: string): boolean {
  return featureFlagManager.getBooleanValue({ name: flagName, defaultValue: "false" });
}

export function constructNode(
  jsonContent: string,
  platform: Platform = Platform.VSCode
): IQTreeNode {
  const jsonObject = JSON.parse(jsonContent);

  // Handle named node references (TypeScript-defined sub-trees)
  if (jsonObject.node) {
    return resolveNodeReference(jsonObject, platform);
  }

  // Handle "group" type (no options)
  if (jsonObject.data.type === "group") {
    const node: IQTreeNode = {
      data: { type: "group", name: jsonObject.data.name },
    };
    if (jsonObject.condition) node.condition = jsonObject.condition;
    if (jsonObject.children) {
      node.children = [];
      for (const child of jsonObject.children) {
        node.children.push(constructNode(JSON.stringify(child), platform));
      }
    }
    return node;
  }

  const data: SingleSelectQuestion = {
    type: "singleSelect",
    name: jsonObject.data.name,
    title: jsonObject.data.title ? getLocalizedString(jsonObject.data.title) : "",
    placeholder: jsonObject.data.placeholder
      ? getLocalizedString(jsonObject.data.placeholder)
      : undefined,
    skipSingleOption: jsonObject.data.skipSingleOption,
    forgetLastValue: jsonObject.data.forgetLastValue,
    staticOptions: [],
    onDidSelection:
      jsonObject.data.name === QuestionNames.WithPlugin ? setTemplateNameAndGC : setTemplateName,
  };

  if (jsonObject.data.options) {
    for (const option of jsonObject.data.options) {
      if (option.featureFlag && !isFeatureEnabled(option.featureFlag)) {
        continue;
      }
      const optionItem: OptionItem = {
        id: option.id as string,
        label:
          option.icon && platform === Platform.VSCode
            ? `${String(option.icon)} ${getLocalizedString(option.label as string)}`
            : getLocalizedString(option.label as string),
        detail: option.detail ? getLocalizedString(option.detail as string) : undefined,
        data: option.data as string,
      };
      if (option.groupName) {
        optionItem.groupName = getLocalizedString(option.groupName);
      }
      (data.staticOptions as OptionItem[]).push(optionItem);
    }
  }

  const node: IQTreeNode = { data };
  if (jsonObject.condition) {
    node.condition = jsonObject.condition;
  }
  if (jsonObject.children) {
    node.children = [];
    for (const child of jsonObject.children) {
      node.children.push(constructNode(JSON.stringify(child), platform));
    }
  }
  return node;
}

function resolveNodeReference(
  jsonObject: { node: string; condition?: Record<string, unknown> | undefined },
  platform: Platform
): IQTreeNode {
  let node: IQTreeNode;

  switch (jsonObject.node) {
    // Common shared nodes (from commonNodes.ts)
    case "llmServiceNode":
      return llmServiceNode(jsonObject.condition as any);
    case "apiSpecNode":
      return apiSpecNode(jsonObject.condition as any);
    case "apiSpecWithSearchNode": {
      const n = apiSpecWithSearchNode();
      if (jsonObject.condition) n.condition = jsonObject.condition;
      return n;
    }
    case "azureOpenAINode":
      return azureOpenAINode(jsonObject.condition as any);
    case "foundryNode":
      return foundryNode(jsonObject.condition as any);

    // TypeScript-defined complex nodes (lazy import to avoid circular dependency)
    case "mcpServerTypeNode": {
      // eslint-disable-next-line @typescript-eslint/no-var-requires
      const { MCPServerTypeNode } = require("./vsc/teamsProjectTypeNode");
      node = MCPServerTypeNode();
      break;
    }

    // Simple question nodes
    case "gcNameNode":
      node = { data: GCNameQuestion() };
      break;
    case "gcConnectionIdNode":
      node = { data: GCConnectionIdQuestion() };
      break;
    case "officeAddinFolderNode":
      node = {
        data: {
          type: "folder",
          name: QuestionNames.OfficeAddinFolder,
          title: getLocalizedString(
            "core.createProjectQuestion.DAMetaOS.capability.upgradeProject.projectFolder.title"
          ),
          placeholder: getLocalizedString(
            "core.createProjectQuestion.DAMetaOS.capability.upgradeProject.projectFolder.placeholder"
          ),
        },
      };
      break;
    case "officeAddinImportNode":
      node = {
        data: { type: "group", name: QuestionNames.OfficeAddinImport },
        children: [
          {
            data: {
              type: "folder",
              name: QuestionNames.OfficeAddinFolder,
              title: "Existing add-in project folder",
            },
          },
          {
            data: {
              type: "singleFile",
              name: QuestionNames.OfficeAddinManifest,
              title: "Select import project manifest file",
            },
          },
        ],
      };
      break;

    // Generic custom inputs node — allows wizard JSON to declare arbitrary
    // question sequences (text, singleSelect, singleFile, folder) without
    // needing new TypeScript code. Inputs are stored in inputs._customInputs
    // and automatically injected into the template replace map.
    case "customInputs":
      return buildCustomInputsNode(jsonObject as any, platform);

    default:
      throw new Error(`Unknown node reference: ${jsonObject.node}`);
  }

  if (jsonObject.condition) {
    node.condition = jsonObject.condition;
  }
  return node;
}

interface CustomInputDef {
  type: "text" | "singleSelect" | "singleFile" | "folder";
  name: string;
  title: string;
  placeholder?: string;
  password?: boolean;
  options?: { id: string; label: string; detail?: string; data?: string }[];
  filters?: { [name: string]: string[] };
  condition?: Record<string, unknown>;
}

interface CustomInputsNodeDef {
  node: "customInputs";
  inputs: CustomInputDef[];
  condition?: Record<string, unknown>;
}

/**
 * Build a question tree from a declarative list of custom inputs.
 *
 * JSON example:
 * ```json
 * {
 *   "node": "customInputs",
 *   "inputs": [
 *     { "type": "text", "name": "endpoint", "title": "template.endpoint.title" },
 *     { "type": "text", "name": "api-key", "title": "template.apiKey.title", "password": true },
 *     { "type": "singleSelect", "name": "auth-type", "title": "template.auth.title",
 *       "options": [{ "id": "none", "label": "None" }, { "id": "key", "label": "API Key" }] },
 *     { "type": "text", "name": "key-value", "title": "template.keyValue.title",
 *       "condition": { "equals": "key" } }
 *   ]
 * }
 * ```
 *
 * Each input becomes a question node. Conditional inputs become children of the
 * previous singleSelect node. User answers are stored in `inputs._customInputs`
 * for automatic injection into the template replace map.
 */
function buildCustomInputsNode(def: CustomInputsNodeDef, _platform: Platform): IQTreeNode {
  const customInputPrefix = "_custom_";

  function buildQuestionData(input: CustomInputDef): Question {
    const title = getLocalizedString(input.title) || input.title;
    const placeholder = input.placeholder
      ? getLocalizedString(input.placeholder) || input.placeholder
      : undefined;

    switch (input.type) {
      case "text":
        return {
          type: "text",
          name: customInputPrefix + input.name,
          title,
          placeholder,
          password: input.password,
        };
      case "singleSelect": {
        const staticOptions: OptionItem[] = (input.options ?? []).map((opt) => ({
          id: opt.id,
          label: getLocalizedString(opt.label) || opt.label,
          detail: opt.detail ? getLocalizedString(opt.detail) || opt.detail : undefined,
          data: opt.data,
        }));
        return {
          type: "singleSelect",
          name: customInputPrefix + input.name,
          title,
          placeholder,
          staticOptions,
        } as SingleSelectQuestion;
      }
      case "singleFile":
        return {
          type: "singleFile",
          name: customInputPrefix + input.name,
          title,
          filters: input.filters,
        };
      case "folder":
        return {
          type: "folder",
          name: customInputPrefix + input.name,
          title,
          placeholder,
        };
      default:
        return { type: "text", name: customInputPrefix + input.name, title };
    }
  }

  // Separate top-level inputs from conditional children
  // Conditional inputs (with condition) attach as children of the most recent singleSelect
  const rootNodes: IQTreeNode[] = [];
  let lastSelectNode: IQTreeNode | undefined;

  for (const input of def.inputs) {
    const questionNode: IQTreeNode = { data: buildQuestionData(input) };

    if (input.condition && lastSelectNode) {
      questionNode.condition = input.condition;
      if (!lastSelectNode.children) lastSelectNode.children = [];
      lastSelectNode.children.push(questionNode);
    } else {
      rootNodes.push(questionNode);
      if (input.type === "singleSelect") {
        lastSelectNode = questionNode;
      }
    }
  }

  // Wrap in a group node, chain as nested children so they appear sequentially
  function chainNodes(nodes: IQTreeNode[]): IQTreeNode {
    if (nodes.length === 0) {
      return { data: { type: "group", name: "custom-inputs" } };
    }
    if (nodes.length === 1) {
      return nodes[0];
    }
    // Nest: first node's children include the rest chained
    const first = nodes[0];
    const rest = chainNodes(nodes.slice(1));
    if (!first.children) first.children = [];
    first.children.push(rest);
    return first;
  }

  // Add an onDidSelection callback to store custom input values
  for (const node of rootNodes) {
    if (node.data && "type" in node.data) {
      const originalName = (node.data as any).name as string;
      if (originalName.startsWith(customInputPrefix)) {
        const cleanName = originalName.slice(customInputPrefix.length);
        const origOnDidSelection = (node.data as any).onDidSelection;
        (node.data as any).onDidSelection = (selected: string | OptionItem, inputs: Inputs) => {
          if (!inputs._customInputs) inputs._customInputs = {};
          const value = typeof selected === "string" ? selected : selected.id;
          inputs._customInputs[cleanName] = value;
          if (origOnDidSelection) origOnDidSelection(selected, inputs);
        };
      }
    }
  }

  const result = chainNodes(rootNodes);
  if (def.condition) {
    result.condition = def.condition;
  }
  return result;
}
