// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Xiaofu Huang <xiaofhua@microsoft.com>
 */
import path from "path";
import * as vscode from "vscode";
import { err, FxError, ok, Result, Stage, Void } from "@microsoft/teamsfx-api";
import { Correlator, TaskDefaultValue } from "@microsoft/teamsfx-core";
import { workspaceUri } from "../../globalVariables";
import { runCommand } from "../../handlers/sharedOpts";
import { TelemetryEvent, TelemetryProperty } from "../../telemetry/extTelemetryEvents";
import { getLocalDebugSession } from "../common/localDebugSession";
import { localTelemetryReporter, maskValue } from "../localTelemetryReporter";
import { BaseTaskTerminal } from "./baseTaskTerminal";
import { getSystemInputs } from "../../utils/systemEnvUtils";

interface LifecycleArgs {
  template?: string;
  env?: string;
}

export class LifecycleTaskTerminal extends BaseTaskTerminal {
  private readonly args: LifecycleArgs;

  constructor(
    taskDefinition: vscode.TaskDefinition,
    private stage: Stage.provision | Stage.deploy
  ) {
    super(taskDefinition);
    this.args = taskDefinition.args as LifecycleArgs;
  }

  do(): Promise<Result<Void, FxError>> {
    const telemetryProperties = {
      [TelemetryProperty.DebugTaskId]: this.taskTerminalId,
      [TelemetryProperty.DebugTaskArgs]: JSON.stringify({
        template: maskValue(this.args.template),
        env: maskValue(this.args.env, [TaskDefaultValue.env]),
      }),
      [TelemetryProperty.DebugLifecycle]: this.stage,
    };

    return Correlator.runWithId(getLocalDebugSession().id, () =>
      localTelemetryReporter.runWithTelemetryProperties(
        TelemetryEvent.DebugLifecycleTask,
        telemetryProperties,
        () => this._do()
      )
    );
  }

  stop(error?: any): Promise<void> {
    return super.stop(error, false);
  }

  private async _do(): Promise<Result<Void, FxError>> {
    if (!this.args?.env) {
      throw BaseTaskTerminal.taskDefinitionError("env");
    }

    const inputs = getSystemInputs();
    inputs.env = this.args.env;
    inputs.isLocalDebug = true;
    if (this.args.template) {
      inputs.workflowFilePath = path.resolve(
        workspaceUri?.fsPath ?? "",
        BaseTaskTerminal.resolveTeamsFxVariables(this.args.template)
      );
    }

    const res = await runCommand(this.stage, inputs);
    return res.isErr() ? err(res.error) : ok(Void);
  }
}
