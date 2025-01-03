// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { Protocol } from "devtools-protocol";
import * as vscode from "vscode";
import { ANSIColors } from "../debug/common/debugConstants";
import { CopilotDebugLog } from "./copilotDebugLogOutput";
import { VS_CODE_UI } from "../qm/vsc_ui";
interface BotTextMessage {
  messageType: string | undefined;
  text: string;
  createdAt: string;
}

export class WebSocketEventHandler {
  static handleEvent(response: Protocol.Network.WebSocketFrame): number {
    let num = 0;
    if (!this.isWebSocketDataRelevant(response)) {
      return num;
    }

    try {
      // logger.info(`Get WebSocket response from: ${url}`);
      const objects = this.splitObjects(response);
      for (const object of objects) {
        const parsedObject = JSON.parse(object) as { item: { messages: BotTextMessage[] } };
        if (parsedObject.item && parsedObject.item.messages) {
          const botTextMessages = this.selectBotTextMessages(parsedObject);
          for (const botTextMessage of botTextMessages) {
            this.convertBotMessageToChannelOutput(botTextMessage);
            num++;
          }
        }
      }
    } catch (error) {
      void VS_CODE_UI.showMessage(
        "error",
        `Error parsing response, ${(error as Error).message}`,
        false
      );
      vscode.debug.activeDebugConsole.appendLine(
        `${ANSIColors.RED} (×) Error: ${ANSIColors.WHITE} Error parsing response: ${
          (error as Error).message
        }`
      );
    }
    return num;
  }

  // only type 2 messages contain developer logs
  static isWebSocketDataRelevant(response: Protocol.Network.WebSocketFrame) {
    return response.payloadData.startsWith('{"type":2');
  }

  static splitObjects(response: Protocol.Network.WebSocketFrame): string[] {
    const objectSeparator = "\x1e"; // ASCII record separator
    return response.payloadData
      .split(objectSeparator)
      .filter((object: string) => object.length > 0);
  }

  static selectBotTextMessages(object: { item: { messages: BotTextMessage[] } }): BotTextMessage[] {
    return object.item.messages.filter((message) => message.messageType === "DeveloperLogs");
  }

  static convertBotMessageToChannelOutput(botTextMessage: BotTextMessage): void {
    new CopilotDebugLog(botTextMessage.text).write();
  }

  static convertBotMessageToChannelOutputJson(botTextMessage: BotTextMessage): CopilotDebugLog {
    return new CopilotDebugLog(this.prettyPrintJson(botTextMessage.text));
  }

  static prettyPrintJson(jsonText: string): string {
    return JSON.stringify(JSON.parse(jsonText), null, 2);
  }
}
