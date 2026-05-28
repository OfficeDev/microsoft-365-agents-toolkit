// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import {
  commands,
  Diagnostic,
  ExtensionContext,
  extensions,
  Uri,
  Range,
  Position,
  languages,
} from "vscode";

import {
  err,
  FxError,
  IDiagnosticInfo,
  InputResult,
  ok,
  Result,
  SingleFileOrInputConfig,
} from "@microsoft/teamsfx-api";
import {
  assembleError,
  isValidHttpUrl,
  loadingDefaultPlaceholder,
  loadingOptionsPlaceholder,
  maskSecret,
} from "@microsoft/teamsfx-core";
import { InternalUIError, Localizer, sleep, VSCodeUI } from "@microsoft/vscode-ui";
import * as packageJson from "../../package.json";
import { TerminalName } from "../constants";
import { ExtTelemetry } from "../telemetry/extTelemetry";
import { getDefaultString, localize } from "../utils/localizeUtils";
import {
  SelectFileOrInputResultType,
  TelemetryEvent,
  TelemetryProperty,
} from "../telemetry/extTelemetryEvents";
import { diagnosticCollection, setDiagnosticCollection } from "../globalVariables";
import { featureFlagManager } from "@microsoft/teamsfx-core";
import { FeatureFlags } from "@microsoft/teamsfx-core";

export class TTKLocalizer implements Localizer {
  invalidInputErrorMessage?(name: string, reason: string): string {
    throw new Error("Method not implemented.");
  }
  invalidInputDisplayMessage?(name: string, reason: string): string {
    throw new Error("Method not implemented.");
  }
  missingInputErrorMessage?(name: string): string {
    throw new Error("Method not implemented.");
  }
  missingInputDisplayMessage?(name: string): string {
    throw new Error("Method not implemented.");
  }
  loadingOptionsPlaceholder(): string {
    return loadingOptionsPlaceholder();
  }
  loadingDefaultPlaceholder(): string {
    return loadingDefaultPlaceholder();
  }
  loadingOptionsTimeoutMessage(): string {
    return "loading options timeout";
  }
  multiSelectKeyboardPlaceholder(): string {
    return localize("teamstoolkit.qm.multiSelectKeyboard");
  }
  defaultFolder(): string {
    return localize("teamstoolkit.qm.defaultFolder");
  }
  browse(): string {
    return localize("teamstoolkit.qm.browse");
  }
  emptyOptionErrorMessage(): string {
    return getDefaultString("teamstoolkit.qm.emptySelection");
  }
  emptyOptionErrorDisplayMessage(): string {
    return localize("teamstoolkit.qm.emptySelection");
  }
  cancelErrorMessage(): string {
    return "User canceled.";
  }
  cancelErrorDisplayMessage(): string {
    return "User canceled.";
  }
  internalErrorDisplayMessage(action: string): string {
    return "VS Code failed to operate: " + action;
  }
  internalErrorMessage(action: string): string {
    return "VS Code failed to operate: " + action;
  }
  commandTimeoutErrorMessage(command: string): string {
    return "Execute command timeout: " + maskSecret(command, { replace: "***" });
  }
  commandTimeoutErrorDisplayMessage(command: string): string {
    return "Execute command timeout: " + maskSecret(command, { replace: "***" });
  }
  commandExecutionErrorMessage(command: string): string {
    return "Execute command failure: " + maskSecret(command, { replace: "***" });
  }
  commandExecutionErrorDisplayMessage(command: string): string {
    return "Execute command failure: " + maskSecret(command, { replace: "***" });
  }
}

export const ttkLocalizer = new TTKLocalizer();
export let VS_CODE_UI: VsCodeUI;

export class VsCodeUI extends VSCodeUI {
  context: ExtensionContext;
  constructor(context: ExtensionContext) {
    super(TerminalName, assembleError, ttkLocalizer);
    this.context = context;
  }
  async reload(): Promise<Result<boolean, FxError>> {
    // The following code only fixes the bug that cause telemetry event lost for projectMigrator().
    // When this reload() function has more users, they may need to dispose() more resources that allocated in activate().
    const extension = extensions.getExtension(`${packageJson.publisher}.${packageJson.name}`);
    if (!extension?.isActive) {
      // When our extension is not activated, we can determine this is in the vscode extension activate() context.
      // Since we are not activated yet, vscode will not deactivate() and dispose() our resourses (which have been allocated in activate()).
      // This may cause resource leaks.For example, buffered events in TelemetryReporter is not sent.
      // So manually dispose them.
      await ExtTelemetry.reporter?.dispose();
    }

    // wait 2 seconds before reloading.
    await sleep(2000);
    const success = await commands.executeCommand("workbench.action.reloadWindow");
    if (success) {
      return ok(success as boolean);
    } else {
      return err(
        new InternalUIError(
          ttkLocalizer.internalErrorMessage(
            `commands.executeCommand("workbench.action.reloadWindow")`
          ),
          ttkLocalizer.internalErrorDisplayMessage(
            `commands.executeCommand("workbench.action.reloadWindow")`
          )
        )
      );
    }
  }

  /**
   * override selectFileOrInput() to send telemetry event
   */
  async selectFileOrInput(
    config: SingleFileOrInputConfig
  ): Promise<Result<InputResult<string>, FxError>> {
    const res = await super.selectFileOrInput(config);
    if (res.isOk()) {
      if (res.value.type === "success") {
        const value = res.value.result as string;
        if (isValidHttpUrl(value)) {
          ExtTelemetry.sendTelemetryEvent(TelemetryEvent.selectFileOrInputResultType, {
            [TelemetryProperty.SelectedOption]: SelectFileOrInputResultType.Input,
          });
        } else {
          ExtTelemetry.sendTelemetryEvent(TelemetryEvent.selectFileOrInputResultType, {
            [TelemetryProperty.SelectedOption]: SelectFileOrInputResultType.LocalFile,
          });
        }
      }
    }
    return res;
  }

  showDiagnosticInfo(diagnostics: IDiagnosticInfo[]): void {
    if (!featureFlagManager.getBooleanValue(FeatureFlags.ShowDiagnostics)) {
      return;
    }
    if (!diagnosticCollection) {
      const collection = languages.createDiagnosticCollection("teamstoolkit");
      setDiagnosticCollection(collection);
    } else {
      diagnosticCollection.clear();
    }
    const diagnosticMap: Map<string, Diagnostic[]> = new Map();
    for (const diagnostic of diagnostics) {
      let diagnosticsOfFile = diagnosticMap.get(diagnostic.filePath);
      if (!diagnosticsOfFile) {
        diagnosticsOfFile = [];
        diagnosticMap.set(diagnostic.filePath, diagnosticsOfFile);
      }

      const diagnosticInVSC = new Diagnostic(
        new Range(
          new Position(diagnostic.startLine, diagnostic.startIndex),
          new Position(diagnostic.endLine, diagnostic.endIndex)
        ),
        diagnostic.message,
        diagnostic.severity
      );
      if (diagnostic.code) {
        diagnosticInVSC.code = {
          value: diagnostic.code.value,
          target: Uri.parse(diagnostic.code.link),
        };
      }
      diagnosticsOfFile.push(diagnosticInVSC);
    }
    diagnosticMap.forEach((diags, filePath) => {
      diagnosticCollection.set(Uri.file(filePath), diags);
    });
  }
}

export function initVSCodeUI(context: ExtensionContext) {
  VS_CODE_UI = new VsCodeUI(context);
}
