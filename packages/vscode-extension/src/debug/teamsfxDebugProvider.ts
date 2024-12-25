// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import path from "path";
import * as vscode from "vscode";

import {
  AppStudioScopes,
  Correlator,
  environmentNameManager,
  envUtil,
  featureFlagManager,
  FeatureFlags,
  Hub,
  isValidProject,
  isValidProjectV3,
  MissingEnvironmentVariablesError,
} from "@microsoft/teamsfx-core";

import VsCodeLogInstance from "../commonlib/log";
import M365TokenInstance from "../commonlib/m365Login";
import { ExtensionSource } from "../error/error";
import { showError } from "../error/common";
import { core } from "../globalVariables";
import { TelemetryEvent, TelemetryProperty } from "../telemetry/extTelemetryEvents";
import { getLocalDebugSessionId, endLocalDebugSession } from "./common/localDebugSession";
import {
  accountHintPlaceholder,
  Host,
  m365AppIdEnv,
  sideloadingDisplayMessages,
} from "./common/debugConstants";
import { localTelemetryReporter, sendDebugAllEvent } from "./localTelemetryReporter";
import { terminateAllRunningTeamsfxTasks } from "./teamsfxTaskHandler";
import { triggerV3Migration } from "../utils/migrationUtils";
import { getSystemInputs } from "../utils/systemEnvUtils";
import { TeamsfxDebugConfiguration } from "./common/teamsfxDebugConfiguration";
import { AgentHintData } from "./common/types";

export class TeamsfxDebugProvider implements vscode.DebugConfigurationProvider {
  public async resolveDebugConfiguration?(
    folder: vscode.WorkspaceFolder | undefined,
    debugConfiguration: TeamsfxDebugConfiguration,
    token?: vscode.CancellationToken
  ): Promise<vscode.DebugConfiguration | undefined> {
    return await Correlator.runWithId(
      getLocalDebugSessionId(),
      this._resolveDebugConfiguration,
      folder,
      debugConfiguration,
      token
    );
  }

  private async _resolveDebugConfiguration(
    folder: vscode.WorkspaceFolder | undefined,
    debugConfiguration: TeamsfxDebugConfiguration,
    token?: vscode.CancellationToken
  ): Promise<vscode.DebugConfiguration | undefined> {
    let telemetryIsRemote: boolean | undefined = undefined;
    try {
      if (!folder) {
        return debugConfiguration;
      }

      if (typeof debugConfiguration.url !== "string") {
        return debugConfiguration;
      }

      if (!isValidProject(folder.uri.fsPath)) {
        return debugConfiguration;
      }

      // migrate to v3
      if (!isValidProjectV3(folder.uri.fsPath)) {
        await triggerV3Migration();
        return debugConfiguration;
      }

      // resolve env
      let url: string = debugConfiguration.url;
      const host = new URL(url).host;
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

      // NOTE: handle the case that msedge/chrome will be resolved twice
      env = env || debugConfiguration.teamsfxEnv;

      const isLocal =
        (env && !environmentNameManager.isRemoteEnvironment(env)) ||
        !debugConfiguration.name.startsWith("Launch Remote");
      telemetryIsRemote = !isLocal;

      // Put env and hub in `debugConfiguration` so debug handlers can retrieve it and send telemetry
      debugConfiguration.teamsfxIsRemote = !isLocal;
      debugConfiguration.teamsfxEnv = env;
      if (host === Host.teams) {
        debugConfiguration.teamsfxHub = Hub.teams;
      } else if (host === Host.outlook) {
        debugConfiguration.teamsfxHub = Hub.outlook;
      } else if (host === Host.office) {
        debugConfiguration.teamsfxHub = Hub.office;
      }

      // Attach correlation-id to DebugConfiguration so concurrent debug sessions are correctly handled in this stage.
      // For backend and bot debug sessions, debugConfiguration.url is undefined so we need to set correlation id early.
      debugConfiguration.teamsfxCorrelationId = getLocalDebugSessionId();

      const result = await localTelemetryReporter.runWithTelemetryExceptionProperties(
        TelemetryEvent.DebugProviderResolveDebugConfiguration,
        {
          [TelemetryProperty.DebugRemote]: (!isLocal).toString(),
          [TelemetryProperty.Hub]: debugConfiguration.teamsfxHub?.toString() ?? "undefined",
        },
        async () => {
          if (debugConfiguration.timeout === undefined) {
            debugConfiguration.timeout = 20000;
          }

          if (env && matchResult) {
            // replace environment variable
            const envRes = await envUtil.readEnv(folder.uri.fsPath, env, false, true);
            if (envRes.isErr()) {
              throw envRes.error;
            }
            const key = matchResult[matchResult.length - 1];
            if (!envRes.value[key]) {
              throw new MissingEnvironmentVariablesError(
                ExtensionSource,
                key,
                path.normalize(path.join(folder.uri.fsPath, ".vscode", "launch.json")),
                "https://aka.ms/teamsfx-tasks"
              );
            }
            url = url.replace(matchResult[0], envRes.value[key]);
          }

          // replace ${account-hint}
          if (url.includes(accountHintPlaceholder)) {
            url = url.replace(
              accountHintPlaceholder,
              await generateAccountHint(host === Host.teams)
            );
          }

          const agentHintMatch = /\${([^:]+):agent-hint}|\${agent-hint}/.exec(url);
          if (agentHintMatch) {
            url = url.replace(
              agentHintMatch[0],
              await generateAgentHint(folder.uri.fsPath, agentHintMatch[1])
            );
          }

          return url;
        }
      );

      if (result === undefined) {
        return undefined;
      }
      debugConfiguration.url = result;

      // NOTE: handle the case that msedge/chrome will be resolved twice
      if (!debugConfiguration.teamsfxResolved) {
        VsCodeLogInstance.info(
          sideloadingDisplayMessages.title(debugConfiguration.teamsfxHub ?? Hub.teams)
        );
        VsCodeLogInstance.outputChannel.appendLine("");
        VsCodeLogInstance.outputChannel.appendLine(
          sideloadingDisplayMessages.sideloadingUrlMessage(
            debugConfiguration.teamsfxHub ?? Hub.teams,
            debugConfiguration.url
          )
        );
        if (isLocal) {
          VsCodeLogInstance.outputChannel.appendLine("");
          VsCodeLogInstance.outputChannel.appendLine(
            sideloadingDisplayMessages.hotReloadingMessage
          );
        }
      }
      debugConfiguration.teamsfxResolved = true;
    } catch (error: any) {
      void showError(error);
      await terminateAllRunningTeamsfxTasks();
      await vscode.debug.stopDebugging();
      // not for undefined
      if (telemetryIsRemote === false) {
        await sendDebugAllEvent(error);
      }
      endLocalDebugSession();
    }
    return debugConfiguration;
  }
}

async function generateAccountHint(includeTenantId = true): Promise<string> {
  let tenantId: string | undefined = undefined;
  let loginHint: string | undefined = undefined;
  const accountInfo = M365TokenInstance.getCachedAccountInfo();
  if (accountInfo !== undefined) {
    tenantId = accountInfo.tenantId;
    loginHint = accountInfo.username;
  } else {
    try {
      const tokenObjectRes = await M365TokenInstance.getStatus({ scopes: AppStudioScopes });
      const tokenObject = tokenObjectRes.isOk() ? tokenObjectRes.value.accountInfo : undefined;
      if (tokenObject) {
        // user signed in
        tenantId = tokenObject.tid as string;
        loginHint = (tokenObject.upn as string) ?? (tokenObject.unique_name as string);
      } else {
        // no signed user
        loginHint = "login_your_m365_account"; // a workaround that user has the chance to login
      }
    } catch {
      // ignore error
    }
  }
  if (includeTenantId && tenantId) {
    if (featureFlagManager.getBooleanValue(FeatureFlags.MultiTenant)) {
      return loginHint
        ? `tenantId=${tenantId}&appTenantId=${tenantId}&login_hint=${loginHint}`
        : "";
    } else {
      return loginHint ? `appTenantId=${tenantId}&login_hint=${loginHint}` : "";
    }
  } else {
    return loginHint ? `login_hint=${loginHint}` : "";
  }
}

async function generateAgentHint(projectPath: string, env: string | undefined): Promise<string> {
  if (!env) {
    env = environmentNameManager.getDefaultEnvName();
  }
  const envRes = await envUtil.readEnv(projectPath, env, false, true);
  if (envRes.isErr()) {
    throw envRes.error;
  }
  if (!envRes.value[m365AppIdEnv]) {
    throw new MissingEnvironmentVariablesError(
      ExtensionSource,
      m365AppIdEnv,
      path.normalize(path.join(projectPath, ".vscode", "launch.json")),
      "https://aka.ms/teamsfx-tasks"
    );
  }
  const id = envRes.value[m365AppIdEnv];
  const clickTimestamp = new Date().toLocaleString();
  const agentHintJson: AgentHintData = {
    id,
    scenario: "launchcopilotextension",
    properties: { clickTimestamp },
    version: 1,
  };
  const base64 = Buffer.from(JSON.stringify(agentHintJson)).toString("base64");
  return base64;
}
