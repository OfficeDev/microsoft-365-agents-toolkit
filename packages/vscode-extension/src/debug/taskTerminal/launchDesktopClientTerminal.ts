// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import * as vscode from "vscode";
import * as util from "util";
import { err, FxError, ok, Result, UserError, Void } from "@microsoft/teamsfx-api";
import { BaseTaskTerminal } from "./baseTaskTerminal";
import {
  AppStudioScopes,
  Correlator,
  envUtil,
  GraphScopes,
  isSovereignHigh,
  MissingEnvironmentVariablesError,
  UserCancelError,
} from "@microsoft/teamsfx-core";
import { localTelemetryReporter, maskValue } from "../localTelemetryReporter";
import { getLocalDebugSession } from "../common/localDebugSession";
import { TelemetryEvent, TelemetryProperty } from "../../telemetry/extTelemetryEvents";
import { ExtensionErrors, ExtensionSource } from "../../error/error";
import { getDefaultString, localize } from "../../utils/localizeUtils";
import { Host, openTerminalDisplayMessage, openTerminalMessage } from "../common/debugConstants";
import { getSystemInputs } from "../../utils/systemEnvUtils";
import { core, tools } from "../../globalVariables";
import path from "path";
import { dotenvUtil } from "@microsoft/teamsfx-core/build/component/utils/envUtil";
import * as fs from "fs";

const showDebugDesktopClientWizard = "SHOW_DEBUG_DESKTOP_CLIENT_WIZARD";

export class LaunchDesktopClientTerminal extends BaseTaskTerminal {
  private readonly url?: string;

  constructor(taskDefinition: vscode.TaskDefinition) {
    super(taskDefinition);
    const args: unknown = taskDefinition.args;
    if (args && typeof args === "object" && "url" in args && typeof args.url === "string") {
      this.url = args.url;
    }
  }

  do(): Promise<Result<Void, FxError>> {
    return Correlator.runWithId(getLocalDebugSession().id, () =>
      localTelemetryReporter.runWithTelemetryProperties(
        TelemetryEvent.LaunchDesktopClientTask,
        {
          [TelemetryProperty.DebugTaskId]: this.taskTerminalId,
          [TelemetryProperty.DebugTaskArgs]: JSON.stringify({
            env: maskValue(this.url),
          }),
        },
        () => this._do()
      )
    );
  }

  private async _do(): Promise<Result<Void, FxError>> {
    if (!vscode.workspace.isTrusted) {
      return err(
        new UserError(
          ExtensionSource,
          ExtensionErrors.UntrustedWorkspaceError,
          getDefaultString("teamstoolkit.localDebug.untrustedWorkspaceError"),
          localize("teamstoolkit.localDebug.untrustedWorkspaceError")
        )
      );
    }

    let url = this.url;
    if (!url?.trim()) {
      return err(BaseTaskTerminal.taskDefinitionError("args.url"));
    }

    const inputs = getSystemInputs();
    const configPath = inputs.projectPath! + "/.localConfigs";
    if (!fs.existsSync(configPath)) {
      fs.writeFileSync(configPath, "");
    }
    const config = dotenvUtil.deserialize(fs.readFileSync(configPath, "utf-8"));
    const loginInfo = await tools.tokenProvider.m365TokenProvider.getStatus({
      scopes: isSovereignHigh() ? GraphScopes : AppStudioScopes(),
    });
    const readMore = `${localize("teamstoolkit.common.readMore")}`;
    if (loginInfo.isOk() && loginInfo.value.status === "SignedIn") {
      const accountInfo = await tools.tokenProvider.m365TokenProvider.getJsonObject({
        scopes: isSovereignHigh() ? GraphScopes : AppStudioScopes(),
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

    let target: URL;
    try {
      target = new URL(`https://${url}`);
    } catch {
      return err(BaseTaskTerminal.taskDefinitionError("args.url"));
    }

    // Require a bare authority, including when URL parsing would normalize it.
    const authority = url.split(/[/?#]/, 1)[0].toLowerCase();
    if (authority !== Host.teams || target.hostname !== Host.teams) {
      return err(BaseTaskTerminal.taskDefinitionError("args.url"));
    }

    const scheme =
      process.platform === "win32" || process.platform === "darwin" ? "msteams" : "https";
    return await this.openDesktopUrl(vscode.Uri.parse(target.href).with({ scheme }));
  }

  private async openDesktopUrl(uri: vscode.Uri): Promise<Result<Void, FxError>> {
    try {
      if (await vscode.env.openExternal(uri)) {
        return ok(Void);
      }
    } catch {
      // Opener errors can include the full URI. Report a localized error without it.
    }

    const isDesktop = uri.scheme === "msteams";
    const name = isDesktop
      ? ExtensionErrors.LaunchTeamsDesktopClientError
      : ExtensionErrors.LaunchTeamsWebClientError;
    const key = isDesktop
      ? "teamstoolkit.localDebug.launchTeamsDesktopClientError"
      : "teamstoolkit.localDebug.launchTeamsWebClientError";
    return err(
      new UserError(
        ExtensionSource,
        name,
        `${getDefaultString(key)} ${openTerminalMessage()}`,
        `${localize(key)} ${openTerminalDisplayMessage()}`
      )
    );
  }
}
