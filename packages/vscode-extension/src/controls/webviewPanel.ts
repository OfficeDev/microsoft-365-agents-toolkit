// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import path from "path";
import * as vscode from "vscode";

import {
  Correlator,
  FeatureFlags,
  SampleConfig,
  featureFlagManager,
  sampleProvider,
} from "@microsoft/teamsfx-core";

import * as extensionPackage from "../../package.json";
import { GlobalKey } from "../constants";
import { TreatmentVariableValue } from "../exp/treatmentVariables";
import * as globalVariables from "../globalVariables";
import { downloadSampleApp } from "../handlers/downloadSample";
import { ExtTelemetry } from "../telemetry/extTelemetry";
import {
  InProductGuideInteraction,
  TelemetryEvent,
  TelemetryProperty,
  TelemetryTriggerFrom,
} from "../telemetry/extTelemetryEvents";
import { localize } from "../utils/localizeUtils";
import { getTriggerFromProperty, isTriggerFromWalkThrough } from "../utils/telemetryUtils";
import { compare } from "../utils/versionUtil";
import { Commands } from "./Commands";
import { PanelType } from "./PanelType";

export class WebviewPanel {
  private static readonly viewType = "react";
  public static currentPanels: WebviewPanel[] = [];

  private panel: vscode.WebviewPanel;
  private panelType: PanelType = PanelType.SampleGallery;
  private disposables: vscode.Disposable[] = [];

  public static createOrShow(panelType: PanelType, args?: any[]) {
    const column = vscode.window.activeTextEditor
      ? vscode.window.activeTextEditor.viewColumn
      : undefined;
    if (
      WebviewPanel.currentPanels &&
      WebviewPanel.currentPanels.findIndex((panel) => panel.panelType === panelType) > -1
    ) {
      WebviewPanel.currentPanels
        .find((panel) => panel.panelType === panelType)!
        .panel.reveal(column);
    } else {
      const isToSide = isTriggerFromWalkThrough(args);
      isToSide
        ? WebviewPanel.currentPanels.push(
            new WebviewPanel(panelType, column || vscode.ViewColumn.Two)
          )
        : WebviewPanel.currentPanels.push(
            new WebviewPanel(panelType, column || vscode.ViewColumn.One)
          );
    }
    // if args empty or undefined, return
    if (!args?.length) {
      return;
    }
    if (panelType == PanelType.SampleGallery && args.length > 1 && typeof args[1] == "string") {
      try {
        const sampleId = args[1];
        const panel = WebviewPanel.currentPanels.find((panel) => panel.panelType === panelType);
        if (panel) {
          void globalVariables.context.globalState.update(
            GlobalKey.SampleGalleryInitialSample,
            sampleId
          );
          ExtTelemetry.sendTelemetryEvent(TelemetryEvent.SelectSample, {
            ...getTriggerFromProperty(args),
            [TelemetryProperty.SampleAppName]: sampleId,
          });
        }
      } catch (e) {}
    }
  }

  private constructor(panelType: PanelType, column: vscode.ViewColumn) {
    this.panelType = panelType;

    // Create and show a new webview panel
    this.panel = vscode.window.createWebviewPanel(
      WebviewPanel.viewType,
      this.getWebpageTitle(panelType),
      column,
      {
        // Enable javascript in the webview
        enableScripts: true,
        retainContextWhenHidden: true,
        localResourceRoots: [
          vscode.Uri.file(path.join(globalVariables.context.extensionPath, "out")),
        ],
      }
    );

    // Listen for when the panel is disposed
    // This happens when the user closes the panel or when the panel is closed programatically
    this.panel.onDidDispose(() => this.dispose(), null, this.disposables);

    this.panel.onDidChangeViewState(
      (e) => {
        const panel = e.webviewPanel;
        if (
          TreatmentVariableValue.inProductDoc &&
          (panelType === PanelType.RespondToCardActions ||
            panelType === PanelType.FunctionBasedNotificationBotReadme ||
            panelType === PanelType.ExpressServerNotificationBotReadme)
        ) {
          ExtTelemetry.sendTelemetryEvent(TelemetryEvent.InteractWithInProductDoc, {
            [TelemetryProperty.TriggerFrom]: TelemetryTriggerFrom.InProductDoc,
            [TelemetryProperty.Interaction]: panel.visible
              ? InProductGuideInteraction.Show
              : InProductGuideInteraction.Hide,
            [TelemetryProperty.Identifier]: panelType,
          });
        } else if (panelType === PanelType.AccountHelp) {
          ExtTelemetry.sendTelemetryEvent(TelemetryEvent.InteractWithInProductDoc, {
            [TelemetryProperty.TriggerFrom]: TelemetryTriggerFrom.AccountHelp,
            [TelemetryProperty.Interaction]: panel.visible
              ? InProductGuideInteraction.Show
              : InProductGuideInteraction.Hide,
            [TelemetryProperty.Identifier]: panelType,
          });
        }
      },
      null,
      globalVariables.context.subscriptions
    );

    // Handle messages from the webview
    this.panel.webview.onDidReceiveMessage(
      async (msg) => {
        switch (msg.command) {
          case Commands.OpenExternalLink:
            void vscode.env.openExternal(vscode.Uri.parse(msg.data));
            break;
          case Commands.CloneSampleApp:
            await Correlator.run(async () => {
              await downloadSampleApp(TelemetryTriggerFrom.Webview, msg.data.appFolder);
            });
            break;
          case Commands.DisplayCommands:
            // eslint-disable-next-line @typescript-eslint/restrict-template-expressions
            await vscode.commands.executeCommand("workbench.action.quickOpen", `>${msg.data}`);
            break;
          case Commands.CreateNewProject:
            await vscode.commands.executeCommand(
              "fx-extension.create",
              TelemetryTriggerFrom.Webview
            );
            break;
          case Commands.SwitchPanel:
            WebviewPanel.createOrShow(msg.data);
            break;
          case Commands.SendTelemetryEvent:
            ExtTelemetry.sendTelemetryEvent(msg.data.eventName, msg.data.properties);
            break;
          case Commands.LoadSampleCollection:
            await this.LoadSampleCollection();
            break;
          case Commands.LoadSampleReadme:
            await this.LoadSampleReadme(msg.data);
            break;
          case Commands.UpgradeToolkit:
            await this.OpenToolkitInExtensionView(msg.data.version);
            break;
          case Commands.StoreData:
            await globalVariables.context.globalState.update(msg.data.key, msg.data.value);
            break;
          case Commands.GetData:
            await this.panel.webview.postMessage({
              message: Commands.GetData,
              data: {
                key: msg.data.key,
                value: globalVariables.context.globalState.get(msg.data.key),
              },
            });
            break;
          case Commands.InvokeTeamsAgent:
            await vscode.commands.executeCommand(
              "fx-extension.invokeChat",
              TelemetryTriggerFrom.Webview
            );
            break;
          default:
            break;
        }
      },
      undefined,
      globalVariables.context.subscriptions
    );

    // Set the webview's initial html content
    this.panel.webview.html = this.getHtmlForWebview(panelType);
    this.panel.iconPath = this.getWebviewPanelIconPath(panelType);
  }

  private async LoadSampleCollection() {
    try {
      await sampleProvider.refreshSampleConfig();
    } catch (e: unknown) {
      await this.panel.webview.postMessage({
        message: Commands.LoadSampleCollection,
        samples: [],
        error: e,
      });
      return;
    }
    const sampleCollection = await sampleProvider.SampleCollection;
    const sampleData = sampleCollection.samples.map((sample) => {
      const extensionVersion = extensionPackage.version;
      let versionComparisonResult = 0;
      if (
        sample.maximumToolkitVersion &&
        compare(extensionVersion, sample.maximumToolkitVersion) > 0
      ) {
        versionComparisonResult = 1;
      }
      if (
        sample.minimumToolkitVersion &&
        compare(extensionVersion, sample.minimumToolkitVersion) < 0
      ) {
        versionComparisonResult = -1;
      }
      return {
        ...sample,
        versionComparisonResult,
      };
    });
    const initialSample = globalVariables.context.globalState.get<string>(
      GlobalKey.SampleGalleryInitialSample,
      ""
    );
    if (this.panel && this.panel.webview) {
      await this.panel.webview.postMessage({
        message: Commands.LoadSampleCollection,
        samples: sampleData,
        initialSample: initialSample,
        filterOptions: sampleCollection.filterOptions,
      });
      if (initialSample != "") {
        // reset initial sample after shown
        await globalVariables.context.globalState.update(GlobalKey.SampleGalleryInitialSample, "");
      }
    }
  }

  private async LoadSampleReadme(sample: SampleConfig) {
    let htmlContent = "";
    try {
      htmlContent = await sampleProvider.getSampleReadmeHtml(sample);
    } catch (e: unknown) {
      await this.panel.webview.postMessage({
        message: Commands.LoadSampleReadme,
        error: e,
        readme: "",
      });
      return;
    }
    if (this.panel && this.panel.webview) {
      let readme = this.replaceRelativeImagePaths(htmlContent, sample);
      readme = this.replaceRelativeMarkdownPaths(readme, sample);
      readme = this.replaceMermaidRelatedContent(readme);
      readme = this.addTabIndex(readme);
      await this.panel.webview.postMessage({
        message: Commands.LoadSampleReadme,
        readme: readme,
      });
    }
  }

  private async OpenToolkitInExtensionView(version: string) {
    // await vscode.commands.executeCommand(
    //   "workbench.extensions.installExtension",
    //   `teamsdevapp.ms-teams-vscode-extension@${version}`
    // );
    await vscode.commands.executeCommand("workbench.extensions.action.checkForUpdates");
  }

  private replaceRelativeImagePaths(htmlContent: string, sample: SampleConfig) {
    const urlInfo = sample.downloadUrlInfo;
    const imageUrl = `https://github.com/${urlInfo.owner}/${urlInfo.repository}/blob/${urlInfo.ref}/${urlInfo.dir}/`;
    const imageRegex = /img\s+src="(?!https:\/\/camo\.githubusercontent\.com\/.)([^"]+)"/gm;
    return htmlContent.replace(imageRegex, `img src="${imageUrl}$1?raw=1"`);
  }

  private replaceMermaidRelatedContent(htmlContent: string): string {
    const mermaidRegex = /<pre lang="mermaid"/gm;
    const loaderRegex = /<span(.*)>\s.*\s*<circle(.*)<\/circle>\s.*<\/path>\s.*\s*<\/span>/gm;
    const loaderRemovedHtmlContent = htmlContent.replace(loaderRegex, "");
    return loaderRemovedHtmlContent.replace(mermaidRegex, `<pre class="mermaid"`);
  }

  private replaceRelativeMarkdownPaths(htmlContent: string, sample: SampleConfig) {
    const markdownRegex = /a\shref="([\.\-_/\\\w]*\.md)"/g;
    const urlInfo = sample.downloadUrlInfo;
    const markdownUrl = `https://github.com/${urlInfo.owner}/${urlInfo.repository}/tree/${urlInfo.ref}/${urlInfo.dir}`;
    const result = htmlContent.replace(markdownRegex, `a href="${markdownUrl}/$1"`);
    return result;
  }

  private addTabIndex(htmlContent: string): string {
    const tabIndexRegex = /<(p|h1|h2|h3|li)/gm;
    return htmlContent.replace(tabIndexRegex, `<$1 tabIndex="0"`);
  }

  private getWebpageTitle(panelType: PanelType): string {
    switch (panelType) {
      case PanelType.SampleGallery:
        return localize("teamstoolkit.webview.samplePageTitle");
      case PanelType.RespondToCardActions:
        return localize("teamstoolkit.guides.cardActionResponse.label");
      case PanelType.AccountHelp:
        return localize("teamstoolkit.webview.accountHelp");
      case PanelType.ExpressServerNotificationBotReadme:
        return localize("teamstoolkit.guides.notificationBot.label");
      case PanelType.FunctionBasedNotificationBotReadme:
        return localize("teamstoolkit.guides.notificationBot.label");
    }
  }

  private getHtmlForWebview(panelType: PanelType) {
    const scriptBasePathOnDisk = vscode.Uri.file(
      path.join(globalVariables.context.extensionPath, "out/")
    );
    const scriptBaseUri = this.panel.webview.asWebviewUri(scriptBasePathOnDisk);

    const scriptPathOnDisk = vscode.Uri.file(
      path.join(globalVariables.context.extensionPath, "out/src", "client.js")
    );
    const scriptUri = this.panel.webview.asWebviewUri(scriptPathOnDisk);
    const codiconsUri = this.panel.webview.asWebviewUri(
      vscode.Uri.joinPath(globalVariables.context.extensionUri, "out", "resource", "codicon.css")
    );
    const stylesheetUri = this.panel.webview.asWebviewUri(
      vscode.Uri.joinPath(globalVariables.context.extensionUri, "out", "resource", "client.css")
    );
    const dompurifyUri = this.panel.webview.asWebviewUri(
      vscode.Uri.joinPath(globalVariables.context.extensionUri, "out", "resource", "purify.min.js")
    );
    const mermaidUri = this.panel.webview.asWebviewUri(
      vscode.Uri.joinPath(globalVariables.context.extensionUri, "out", "resource", "mermaid.min.js")
    );

    const allowChat = featureFlagManager.getBooleanValue(FeatureFlags.ChatParticipantUIEntries);

    // Use a nonce to to only allow specific scripts to be run
    const nonce = this.getNonce();
    return `<!DOCTYPE html>
        <html lang="en">
          <head>
            <meta charset="UTF-8">
            <meta name="viewport" content="width=device-width, initial-scale=1.0">
            <title>ms-teams</title>
            <base href='${scriptBaseUri.toString()}' />
            <link href="${stylesheetUri.toString()}" rel="stylesheet" />
            <link href="${codiconsUri.toString()}" rel="stylesheet" />
          </head>
          <body>
            <div id="root"></div>
            <script>
              const vscode = acquireVsCodeApi();
              const panelType = '${panelType}';
              const shouldShowChat = '${allowChat ? "true" : "false"}';
            </script>
            <script nonce="${nonce}" type="module" src="${scriptUri.toString()}"></script>
            <script nonce="${nonce}" type="text/javascript" src="${dompurifyUri.toString()}"></script>
            <script nonce="${nonce}" type="text/javascript" src="${mermaidUri.toString()}">
            </script>
          </body>
        </html>`;
  }

  private getNonce() {
    let text = "";
    // eslint-disable-next-line
    const possible = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789";
    for (let i = 0; i < 32; i++) {
      text += possible.charAt(Math.floor(Math.random() * possible.length));
    }
    return text;
  }

  private getWebviewPanelIconPath(panelType: PanelType) {
    if (panelType === PanelType.AccountHelp) {
      return vscode.Uri.file(
        path.join(globalVariables.context.extensionPath, "media/itp/m365icon.svg")
      );
    }
    return undefined;
  }

  public dispose() {
    const panelIndex = WebviewPanel.currentPanels.indexOf(this);
    if (TreatmentVariableValue.inProductDoc && this.panelType === PanelType.RespondToCardActions) {
      ExtTelemetry.sendTelemetryEvent(TelemetryEvent.InteractWithInProductDoc, {
        [TelemetryProperty.TriggerFrom]: TelemetryTriggerFrom.InProductDoc,
        [TelemetryProperty.Interaction]: InProductGuideInteraction.Close,
        [TelemetryProperty.Identifier]: this.panelType,
      });
    } else if (this.panelType === PanelType.AccountHelp) {
      ExtTelemetry.sendTelemetryEvent(TelemetryEvent.InteractWithInProductDoc, {
        [TelemetryProperty.TriggerFrom]: TelemetryTriggerFrom.AccountHelp,
        [TelemetryProperty.Interaction]: InProductGuideInteraction.Close,
        [TelemetryProperty.Identifier]: this.panelType,
      });
    }

    WebviewPanel.currentPanels.splice(panelIndex, 1);

    // Clean up our resources
    this.panel.dispose();

    while (this.disposables.length) {
      const x = this.disposables.pop();
      if (x) {
        x.dispose();
      }
    }
  }
}
