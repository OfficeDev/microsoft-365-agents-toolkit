// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import * as cp from "child_process";
import * as vscode from "vscode";
import * as util from "util";
import { err, FxError, ok, Result, UserError, Void } from "@microsoft/teamsfx-api";
import { BaseTaskTerminal } from "./baseTaskTerminal";
import {
  AppStudioScopes,
  Correlator,
  envUtil,
  MissingEnvironmentVariablesError,
  UserCancelError,
} from "@microsoft/teamsfx-core";
import { localTelemetryReporter, maskValue } from "../localTelemetryReporter";
import { getLocalDebugSession } from "../common/localDebugSession";
import { TelemetryEvent, TelemetryProperty } from "../../telemetry/extTelemetryEvents";
import { ExtensionErrors, ExtensionSource } from "../../error/error";
import { getDefaultString, localize } from "../../utils/localizeUtils";
import { openTerminalDisplayMessage, openTerminalMessage } from "../common/debugConstants";
import { getSystemInputs } from "../../utils/systemEnvUtils";
import { core, tools } from "../../globalVariables";
import path from "path";
import { dotenvUtil } from "@microsoft/teamsfx-core/build/component/utils/envUtil";
import * as fs from "fs";

const showDebugDesktopClientWizard = "SHOW_DEBUG_DESKTOP_CLIENT_WIZARD";

interface LaunchDesktopClientArgs {
  url: string;
}

export class LaunchDesktopClientTerminal extends BaseTaskTerminal {
  private readonly args: LaunchDesktopClientArgs;

  constructor(taskDefinition: vscode.TaskDefinition) {
    super(taskDefinition);
    this.args = taskDefinition.args as LaunchDesktopClientArgs;
  }

  do(): Promise<Result<Void, FxError>> {
    return Correlator.runWithId(getLocalDebugSession().id, () =>
      localTelemetryReporter.runWithTelemetryProperties(
        TelemetryEvent.LaunchDesktopClientTask,
        {
          [TelemetryProperty.DebugTaskId]: this.taskTerminalId,
          [TelemetryProperty.DebugTaskArgs]: JSON.stringify({
            env: maskValue(this.args.url),
          }),
        },
        () => this._do()
      )
    );
  }

  private async _do(): Promise<Result<Void, FxError>> {
    const inputs = getSystemInputs();
    const configPath = inputs.projectPath! + "/.localConfigs";
    if (!fs.existsSync(configPath)) {
      fs.writeFileSync(configPath, "");
    }
    const config = dotenvUtil.deserialize(fs.readFileSync(configPath, "utf-8"));
    const loginInfo = await tools.tokenProvider.m365TokenProvider.getStatus({
      scopes: AppStudioScopes,
    });
    const readMore = `${localize("teamstoolkit.common.readMore")}`;
    if (loginInfo.isOk() && loginInfo.value.status === "SignedIn") {
      const accountInfo = await tools.tokenProvider.m365TokenProvider.getJsonObject({
        scopes: AppStudioScopes,
      });
      let username = "";
      if (accountInfo.isOk() && accountInfo.value["unique_name"]) {
        username = " (" + (accountInfo.value["unique_name"] as string) + ")";
      }
      if (config.obj[showDebugDesktopClientWizard] === "false") {
        void vscode.window
          .showWarningMessage(
            util.format(
              localize("teamstoolkit.localDebug.launchTeamsDesktopClientMessage"),
              username
            ),
            { modal: false },
            readMore
          )
          .then((selection) => {
            if (selection === readMore) {
              void vscode.env.openExternal(
                vscode.Uri.parse("https://aka.ms/teamsfx-debug-in-desktop-client")
              );
            }
          });
      } else {
        let userSelected: string | undefined;
        do {
          userSelected = await vscode.window.showInformationMessage(
            util.format(
              localize("teamstoolkit.localDebug.launchTeamsDesktopClientMessage"),
              username
            ),
            { modal: true },
            "Continue",
            `${localize("teamstoolkit.common.readMore")}`
          );
          if (userSelected === readMore) {
            void vscode.env.openExternal(
              vscode.Uri.parse("https://aka.ms/teamsfx-debug-in-desktop-client")
            );
          } else if (userSelected != "Continue") {
            return err(new UserCancelError());
          } else {
            config.obj[showDebugDesktopClientWizard] = "false";
            fs.writeFileSync(configPath, dotenvUtil.serialize(config));
          }
        } while (userSelected === readMore);
      }
    }

    let url: string = this.args.url;
    let env: string | undefined = undefined;

    // match ${{xxx:yyy}}
    let matchResult = /\${{(.+):([A-Za-z0-9_]+)}}/.exec(url);
    if (matchResult) {
      env = matchResult[1];
    }

    if (!env) {
      // match ${{yyy}}
      matchResult = /\${{([A-Za-z0-9_]+)}}/.exec(url);
      if (matchResult) {
        // prompt to select env
        const inputs = getSystemInputs();
        inputs.ignoreEnvInfo = false;
        inputs.ignoreLocalEnv = true;
        const envResult = await core.getSelectedEnv(inputs);
        if (envResult.isErr()) {
          throw envResult.error;
        }
        env = envResult.value;
      }
    }

    if (env && matchResult) {
      // replace environment variable
      const envRes = await envUtil.readEnv(inputs.projectPath!, env, false, true);
      if (envRes.isErr()) {
        throw envRes.error;
      }
      const key = matchResult[matchResult.length - 1];
      if (!envRes.value[key]) {
        throw new MissingEnvironmentVariablesError(
          ExtensionSource,
          key,
          path.normalize(path.join(inputs.projectPath!, "env", ".env." + env)),
          "https://aka.ms/teamsfx-tasks"
        );
      }
      url = url.replace(matchResult[0], envRes.value[key]);
    }

    return await this.openDesktopUrl(url);
  }

  private openDesktopUrl(url: string): Promise<Result<Void, FxError>> {
    return new Promise<Result<Void, FxError>>((resolve) => {
      let childProc;
      if (process.platform === "win32") {
        childProc = cp.exec(`start msteams://${url}`);
        this.writeEmitter.fire(`start msteams://${url}\r\n`);
      } else if (process.platform === "darwin") {
        childProc = cp.exec(`open msteams://${url}`);
        this.writeEmitter.fire(`open msteams://${url}\r\n`);
      } else {
        void vscode.env.openExternal(vscode.Uri.parse("https://" + url));
        childProc = cp.exec(`echo https://${url}`);
      }

      childProc.stdout?.setEncoding("utf-8");
      childProc.stdout?.on("data", (data: string | Buffer) => {
        const line = data.toString().replace(/\n/g, "\r\n");
        this.writeEmitter.fire(line);
      });

      childProc.stderr?.setEncoding("utf-8");
      childProc.stderr?.on("data", (data: string | Buffer) => {
        const line = data.toString().replace(/\n/g, "\r\n");
        this.writeEmitter.fire(line);
      });

      childProc.on("error", (error) => {
        resolve(
          err(
            new UserError(
              ExtensionSource,
              ExtensionErrors.LaunchTeamsDesktopClientError,
              `${getDefaultString("teamstoolkit.localDebug.launchTeamsDesktopClientError")} ${
                error?.message ?? ""
              }  ${openTerminalDisplayMessage()}`,
              `${localize("teamstoolkit.localDebug.launchTeamsDesktopClientError")} ${
                error?.message ?? ""
              }  ${openTerminalDisplayMessage()}`
            )
          )
        );
      });

      childProc.on("close", (code: number) => {
        if (code === 0) {
          resolve(ok(Void));
        } else {
          resolve(
            err(
              new UserError(
                ExtensionSource,
                ExtensionErrors.LaunchTeamsWebClientError,
                util.format(
                  getDefaultString("teamstoolkit.localDebug.launchTeamsDesktopClientStoppedError"),
                  code
                ) +
                  " " +
                  openTerminalMessage(),
                util.format(
                  localize("teamstoolkit.localDebug.launchTeamsDesktopClientStoppedError"),
                  code
                ) +
                  " " +
                  openTerminalDisplayMessage()
              )
            )
          );
        }
      });
    });
  }
}
