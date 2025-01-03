// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { Correlator, maskSecret } from "@microsoft/teamsfx-core";
import * as CDP from "chrome-remote-interface";
import { Protocol } from "devtools-protocol";
import * as uuid from "uuid";
import * as vscode from "vscode";
import logger from "../commonlib/log";
import { connectToExistingBrowserDebugSessionForCopilot } from "../debug/common/debugConstants";
import { VS_CODE_UI } from "../qm/vsc_ui";
import { ExtTelemetry } from "../telemetry/extTelemetry";
import { WebSocketEventHandler } from "./webSocketEventHandler";

export const CDPModule = {
  build: CDP.default,
};

export interface CopilotDebugSession {
  url: string;
  port: number;
  clients: CDP.Client[];
  errors: Error[];
  sessionId: string;
}
// sessions: Map<number, CopilotDebugSession> = new Map();
export class CDPClient {
  url: string;
  port: number;
  sessionId: string;
  name: string;
  client?: CDP.Client;
  errors: Error[] = [];
  messageNumber = 0;

  constructor(url: string, port: number, name: string) {
    this.url = url;
    this.port = port;
    this.name = name;
    this.sessionId = uuid.v4();
  }
  async connectWithBackoff(
    debugPort: number,
    target = "",
    retries = 5,
    delay = 2000
  ): Promise<CDP.Client> {
    let recentError;
    for (let i = 0; i < retries; i++) {
      try {
        await new Promise((resolve) => setTimeout(resolve, delay)); // initial delay
        const client = await CDPModule.build({ port: debugPort, target });
        return client;
      } catch (error) {
        this.errors.push(error);
        recentError = error;
        void vscode.window.showInformationMessage(
          `Attempt ${i + 1} failed. Retrying in ${delay}ms...`
        );
        delay *= 2; // double the delay for the next attempt
      }
    }
    void vscode.window.showErrorMessage("All attempts to connect have failed");
    throw recentError;
  }
  async subscribeToWebSocketEvents(client: CDP.Client): Promise<void> {
    if (this.url.includes("m365.cloud.microsoft/chat")) {
      // only m365.cloud.microsoft need listen to sub target iframe
      void this.launchTeamsChatListener(client);
    } else {
      // Enable the necessary domains
      await client.Network.enable();
      await client.Page.enable();
      client.Network.webSocketFrameReceived(this.webSocketFrameReceivedHandler.bind(this));
      this.client = client;
      logger.info(
        `Connected to copilot debug session: ${this.name}, port: ${this.port}, url: ${this.url}`
      );
    }
  }

  async launchTeamsChatListener(client: CDP.Client, retries = 10, interval = 3000): Promise<void> {
    for (let i = 0; i < retries; ++i) {
      try {
        const res = await this.connectToTargetIframe(client);
        if (res) {
          break;
        }
      } catch (error) {
        this.errors.push(error);
      }
      if (i + 1 >= retries) break;
      await new Promise((resolve) => setTimeout(resolve, interval));
    }
  }

  async connectToTargetIframe(client: CDP.Client): Promise<boolean> {
    const targets = await client.Target.getTargets();
    const iframeTargets = targets.targetInfos.filter(
      ({ type, url }) =>
        type === "iframe" && url.includes("outlook.office.com/hosted/semanticoverview/Users")
    );
    for (const iframeTarget of iframeTargets) {
      const iframeClient = await this.connectWithBackoff(this.port, iframeTarget.targetId);
      if (iframeClient) {
        await client.close();
        await iframeClient.Network.enable();
        await iframeClient.Page.enable();
        iframeClient.Network.webSocketFrameReceived(this.webSocketFrameReceivedHandler.bind(this));
        this.client = iframeClient;
        logger.info(
          `Connected to copilot debug session: ${this.name}, port: ${this.port}, url: ${this.url}`
        );
        return true;
      }
    }
    return false;
  }

  async start() {
    await Correlator.runWithId(this.sessionId, async () => {
      const urlType = isOfficeChatUrl(this.url) ? "office" : "m365";
      ExtTelemetry.sendTelemetryEvent("cdp-client-start", {
        port: `${this.port}`,
        url: urlType,
        name: this.name,
      });
      try {
        const client: CDP.Client = await this.connectWithBackoff(this.port);
        await this.subscribeToWebSocketEvents(client);
        vscode.debug.activeDebugConsole.appendLine(
          connectToExistingBrowserDebugSessionForCopilot.successfulConnectionMessage(this.port)
        );
        ExtTelemetry.sendTelemetryEvent("cdp-client-start-success", {
          port: `${this.port}`,
          url: urlType,
          name: this.name,
          errors: maskSecret(
            this.errors.map((e) => JSON.stringify(e, Object.getOwnPropertyNames(e))).join(",")
          ),
        });
      } catch (error) {
        ExtTelemetry.sendTelemetryErrorEvent("cdp-client-start-fail", error, {
          port: `${this.port}`,
          url: urlType,
          name: this.name,
          errors: maskSecret(
            this.errors.map((e) => JSON.stringify(e, Object.getOwnPropertyNames(e))).join(",")
          ),
        });
      }
    });
  }
  async stop() {
    await Correlator.runWithId(this.sessionId || "", async () => {
      const urlType = isOfficeChatUrl(this.url) ? "office" : "m365";
      await this.client?.close();
      this.client = undefined;
      logger.info(
        `Disconnected to copilot debug session: ${this.name}, port: ${this.port}, url: ${this.url}`
      );
      ExtTelemetry.sendTelemetryEvent("cdp-client-end", {
        port: `${this.port}`,
        url: urlType,
        name: this.name,
        count: `${this.messageNumber}`,
      });
    });
  }
  webSocketFrameReceivedHandler(event: Protocol.Network.WebSocketFrameReceivedEvent) {
    const num = WebSocketEventHandler.handleEvent(event.response);
    this.messageNumber += num;
  }
}

export function isCopilotChatUrl(url: string): boolean {
  const low = url.toLowerCase();
  return low.includes("office.com/chat") || low.includes("m365.cloud.microsoft/chat");
}

export function isOfficeChatUrl(url: string) {
  const low = url.toLowerCase();
  return low.includes("office.com/chat");
}

export function isM365ChatUrl(url: string) {
  const low = url.toLowerCase();
  return low.includes("m365.cloud.microsoft/chat");
}

export function isM365CopilotChatDebugConfiguration(
  configuration: vscode.DebugConfiguration
): number | undefined {
  if (configuration.request !== "launch") return undefined;
  if (!configuration.url) return undefined;
  if (!isCopilotChatUrl(configuration.url)) return undefined;
  if (!configuration.url.includes("developerMode=Basic")) return undefined;
  if (!configuration.runtimeArgs) return undefined;
  const runtimeArgs = configuration.runtimeArgs as string[];
  const portArg = runtimeArgs.find((arg) => arg.startsWith("--remote-debugging-port="));
  if (!portArg) return undefined;
  const port = Number(portArg.substring("--remote-debugging-port=".length));
  if (isNaN(port)) return undefined;
  return port;
}

class CDPClientManager {
  sessions: Map<number, CDPClient> = new Map();
  start(url: string, port: number, name: string): CDPClient {
    const existing = this.sessions.get(port);
    if (existing) {
      // throw new PortsConflictError([port], [port], ExtensionSource);
      void VS_CODE_UI.showMessage(
        "error",
        `Debugging session already exists on this port: ${port}, the existing session will be replaced. Please change the runtimeArgs '--remote-debugging-port' in launch.json and relaunch again.`,
        false
      );
      void existing.stop();
    }
    const session = new CDPClient(url, port, name);
    void session.start();
    this.sessions.set(port, session);
    return session;
  }
  async stop(port: number): Promise<void> {
    const client = this.sessions.get(port);
    if (client) {
      await client.stop();
    }
  }
}

export const cdpClientManager = new CDPClientManager();
