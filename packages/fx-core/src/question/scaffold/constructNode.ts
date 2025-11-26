// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { IQTreeNode, OptionItem, Question } from "@microsoft/teamsfx-api";
import { getLocalizedString } from "../../common/localizeUtils";
import { apiSpecNode, azureOpenAINode, llmServiceNode } from "./commonNodes";
import { setTemplateName } from "./vsc/CapabilityOptions";

export function constructNode(jsonContent: string): IQTreeNode {
  const jsonObject = JSON.parse(jsonContent);
  if (jsonObject.node === "llmServiceNode") {
    return llmServiceNode(jsonObject.condition);
  } else if (jsonObject.node === "apiSpecNode") {
    return apiSpecNode(jsonObject.condition);
  } else if (jsonObject.node === "azureOpenAINode") {
    return azureOpenAINode(jsonObject.condition);
  }
  const data: Question = {
    type: jsonObject.data.type,
    name: jsonObject.data.name,
    title: getLocalizedString(jsonObject.data.title),
    placeholder: getLocalizedString(jsonObject.data.placeholder),
    onDidSelection: setTemplateName,
    staticOptions: [],
  };
  for (const option of jsonObject.data.options) {
    (data.staticOptions as OptionItem[]).push({
      id: option.id,
      label: getLocalizedString(option.label),
      detail: getLocalizedString(option.detail),
      data: option.data,
    });
  }
  const node: IQTreeNode = {
    data: data,
  };
  if (jsonObject.condition) {
    node.condition = jsonObject.condition;
  }
  for (const option of jsonObject.children) {
    const childNode = constructNode(JSON.stringify(option));
    node.children = node.children || [];
    node.children.push(childNode);
  }
  return node;
}
