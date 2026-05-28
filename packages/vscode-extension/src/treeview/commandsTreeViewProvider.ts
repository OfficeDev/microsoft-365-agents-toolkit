// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import * as vscode from "vscode";

import { TreeViewCommand } from "./treeViewCommand";

export class CommandsTreeViewProvider implements vscode.TreeDataProvider<TreeViewCommand> {
  private _onDidChangeTreeData: vscode.EventEmitter<TreeViewCommand | undefined | void> =
    new vscode.EventEmitter<TreeViewCommand | undefined | void>();
  readonly onDidChangeTreeData: vscode.Event<TreeViewCommand | undefined | void> =
    this._onDidChangeTreeData.event;

  private commands: TreeViewCommand[] = [];
  private disposableMap: Map<string, vscode.Disposable> = new Map();

  public constructor(commands: TreeViewCommand[]) {
    this.commands.push(...commands);
  }

  public refresh(): void {
    this._onDidChangeTreeData.fire();
  }

  public getTreeItem(element: TreeViewCommand): vscode.TreeItem {
    const label = element.label
      ? typeof element.label === "string"
        ? element.label
        : element.label.label
      : "";
    const tooltip = element.tooltip
      ? typeof element.tooltip === "string"
        ? element.tooltip
        : element.tooltip.value
      : "";
    element.accessibilityInformation = {
      label: `${label}. ${tooltip}`.trim(),
    };
    return element;
  }

  public getChildren(element?: TreeViewCommand): Thenable<TreeViewCommand[]> {
    if (element && element.children) {
      return Promise.resolve(element.children);
    } else {
      return Promise.resolve(this.commands);
    }
  }

  public getCommands(): TreeViewCommand[] {
    return this.commands;
  }

  dispose() {
    this.disposableMap.forEach((value) => {
      value.dispose();
    });
  }
}
