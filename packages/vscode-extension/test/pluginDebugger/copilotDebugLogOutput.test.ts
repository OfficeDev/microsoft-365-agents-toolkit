import { LogLevel } from "@microsoft/teamsfx-api";
import * as chai from "chai";
import * as sinon from "sinon";
import * as vscode from "vscode";
import { ANSIColors } from "../../src/debug/common/debugConstants";
import * as globalVariables from "../../src/globalVariables";
import {
  CopilotDebugLog,
  logToDebugConsole,
  writeExecutionDetailsToFile,
} from "../../src/pluginDebugger/copilotDebugLogOutput";

describe("copilotDebugLogOutput", () => {
  const sandbox = sinon.createSandbox();
  const message = "log message";
  const fixedDate = new Date("2023-01-01T00:00:00.000Z");
  const logDateString = fixedDate.toJSON();

  beforeEach(() => {
    sandbox.useFakeTimers(fixedDate.getTime());
  });

  afterEach(() => {
    sandbox.restore();
  });

  describe("logToDebugConsole", () => {
    it("should log info messages to the debug console", () => {
      const appendLineStub = sandbox.stub(vscode.debug.activeDebugConsole, "appendLine");
      logToDebugConsole(LogLevel.Info, message);
      chai.assert.isTrue(appendLineStub.calledOnce);
      chai.assert.isTrue(
        appendLineStub.calledWith(
          ANSIColors.WHITE + `[${logDateString}] - ` + ANSIColors.BLUE + `${message}`
        )
      );
    });
    it("should log warning messages to the debug console", () => {
      const appendLineStub = sandbox.stub(vscode.debug.activeDebugConsole, "appendLine");
      logToDebugConsole(LogLevel.Warning, message);
      chai.assert.isTrue(appendLineStub.calledOnce);
      chai.assert.isTrue(
        appendLineStub.calledWith(
          ANSIColors.WHITE + `[${logDateString}] - ` + ANSIColors.YELLOW + `${message}`
        )
      );
    });
    it("should log error messages to the debug console", () => {
      const appendLineStub = sandbox.stub(vscode.debug.activeDebugConsole, "appendLine");
      logToDebugConsole(LogLevel.Error, message);
      chai.assert.isTrue(appendLineStub.calledOnce);
      chai.assert.isTrue(
        appendLineStub.calledWith(
          ANSIColors.WHITE + `[${logDateString}] - ` + ANSIColors.RED + `${message}`
        )
      );
    });
    it("should log debug messages to the debug console", () => {
      const appendLineStub = sandbox.stub(vscode.debug.activeDebugConsole, "appendLine");
      logToDebugConsole(LogLevel.Debug, message);
      chai.assert.isTrue(appendLineStub.calledOnce);
      chai.assert.isTrue(
        appendLineStub.calledWith(
          ANSIColors.WHITE + `[${logDateString}] - ` + ANSIColors.GREEN + `${message}`
        )
      );
    });
    it("should log messages to the debug console", () => {
      const appendLineStub = sandbox.stub(vscode.debug.activeDebugConsole, "appendLine");
      logToDebugConsole(LogLevel.Verbose, message);
      chai.assert.isTrue(appendLineStub.calledOnce);
      chai.assert.isTrue(
        appendLineStub.calledWith(ANSIColors.WHITE + `[${logDateString}] - ${message}`)
      );
    });
  });

  describe("writeExecutionDetailsToFile", () => {
    it("should write function execution details to file", async () => {
      const fs = require("fs");
      const appendFileStub = sandbox.stub(fs, "appendFileSync").resolves();
      writeExecutionDetailsToFile("path/to/log.txt", "log message");
      chai.assert.isTrue(appendFileStub.calledWith("path/to/log.txt", "log message\n"));
    });
  });

  describe("CopilotDebugLog", () => {
    const sandbox = sinon.createSandbox();

    afterEach(() => {
      sandbox.restore();
    });
    it("should parse log JSON and initialize properties", () => {
      const logJson = JSON.stringify({
        enabledPlugins: [{ name: "plugin1", id: "1", version: "1.0" }],
        matchedFunctionCandidates: [
          {
            plugin: { name: "plugin1", id: "1", version: "1.0" },
            functionDisplayName: "function1",
          },
        ],
        functionsSelectedForInvocation: [
          {
            plugin: { name: "plugin1", id: "1", version: "1.0" },
            functionDisplayName: "function1",
          },
        ],
        functionExecutions: [
          {
            function: {
              plugin: { name: "plugin1", id: "1", version: "1.0" },
              functionDisplayName: "function1",
            },
            executionStatus: { requestStatus: 200, responseStatus: 200, responseType: 1 },
            parameters: {},
            requestUri: "http://example.com",
            requestMethod: "GET",
            responseContent: "",
            responseContentType: "",
            errorMessage: "",
          },
        ],
      });
      const copilotDebugLog = new CopilotDebugLog(logJson);
      chai.assert.deepEqual(copilotDebugLog.enabledPlugins, [
        { name: "plugin1", id: "1", version: "1.0" },
      ]);
      chai.assert.deepEqual(copilotDebugLog.matchedFunctionCandidates, [
        { plugin: { name: "plugin1", id: "1", version: "1.0" }, functionDisplayName: "function1" },
      ]);
      chai.assert.deepEqual(copilotDebugLog.functionsSelectedForInvocation, [
        { plugin: { name: "plugin1", id: "1", version: "1.0" }, functionDisplayName: "function1" },
      ]);
      chai.assert.deepEqual(copilotDebugLog.functionExecutions, [
        {
          function: {
            plugin: { name: "plugin1", id: "1", version: "1.0" },
            functionDisplayName: "function1",
          },
          executionStatus: { requestStatus: 200, responseStatus: 200, responseType: 1 },
          parameters: {},
          requestUri: "http://example.com",
          requestMethod: "GET",
          responseContent: "",
          responseContentType: "",
          errorMessage: "",
        },
      ]);
    });

    it("should throw an error if log JSON is invalid", () => {
      const invalidLogJson = "{ invalid json }";
      chai.assert.throws(() => new CopilotDebugLog(invalidLogJson), /Error parsing logAsJson/);
    });

    it("should throw an error if requestUri is invalid", () => {
      const logJson = JSON.stringify({
        functionExecutions: [
          {
            function: {
              plugin: { name: "plugin1", id: "1", version: "1.0" },
              functionDisplayName: "function1",
            },
            executionStatus: { requestStatus: 200, responseStatus: 200, responseType: 1 },
            parameters: {},
            requestUri: "invalid uri",
            requestMethod: "GET",
            responseContent: "",
            responseContentType: "",
            errorMessage: "",
          },
        ],
      });
      chai.assert.throws(
        () => new CopilotDebugLog(logJson),
        /Error creating URL object for requestUri/
      );
    });

    it("write with 0 enabled plugin(s)", () => {
      const logJson = JSON.stringify({
        enabledPlugins: [],
        matchedFunctionCandidates: [],
        functionsSelectedForInvocation: [],
        functionExecutions: [],
      });

      const copilotDebugLog = new CopilotDebugLog(logJson);
      const appendLineStub = sandbox.stub(vscode.debug.activeDebugConsole, "appendLine");
      copilotDebugLog.write();

      chai.assert.isTrue(appendLineStub.calledWith(""));
      chai.assert.isTrue(
        appendLineStub.calledWith(
          `${ANSIColors.WHITE}[${new Date().toJSON()}] - ${ANSIColors.BLUE}0 enabled action(s).`
        )
      );
      chai.assert.isTrue(
        appendLineStub.calledWith(
          `${ANSIColors.RED}(×) Error: ${ANSIColors.WHITE}Enabled plugin: None`
        )
      );
    });

    it("write with plugins enabled", () => {
      const logJson = JSON.stringify({
        enabledPlugins: [{ name: "plugin1", id: "1", version: "1.0" }],
        matchedFunctionCandidates: [
          {
            plugin: { name: "plugin1", id: "1", version: "1.0" },
            functionDisplayName: "function1",
          },
        ],
        functionsSelectedForInvocation: [
          {
            plugin: { name: "plugin1", id: "1", version: "1.0" },
            functionDisplayName: "function1",
          },
        ],
        functionExecutions: [
          {
            function: {
              plugin: { name: "plugin1", id: "1", version: "1.0" },
              functionDisplayName: "function1",
            },
            executionStatus: { requestStatus: 200, responseStatus: 200, responseType: 1 },
            parameters: {},
            requestUri: "http://example.com",
            requestMethod: "GET",
            responseContent: "",
            responseContentType: "",
            errorMessage: "Sample error",
          },
        ],
      });
      const logFilePath = `/path/to/log/Copilot-debug-test.txt`;
      const responseStatus = 200;
      const fs = require("fs");
      const appendFileSyncStub = sandbox.stub(fs, "appendFileSync").resolves();
      sandbox.stub(globalVariables, "defaultExtensionLogPath").value("/path/to/log");
      sandbox.stub(Date.prototype, "toISOString").returns("test");
      const copilotDebugLog = new CopilotDebugLog(logJson);
      const appendLineStub = sandbox.stub(vscode.debug.activeDebugConsole, "appendLine");
      copilotDebugLog.write();
      chai.assert.isTrue(
        appendLineStub.calledWith(
          `${ANSIColors.GREEN}(√) ${ANSIColors.WHITE}Enabled action: ${ANSIColors.MAGENTA}plugin1 ${ANSIColors.GRAY}• version 1.0 • 1`
        )
      );
      chai.assert.isTrue(
        appendLineStub.calledWith(
          `${ANSIColors.GREEN}   (√) ${ANSIColors.WHITE}Matched functions: ${ANSIColors.MAGENTA}function1`
        )
      );
      chai.assert.isTrue(
        appendLineStub.calledWith(
          `${ANSIColors.GREEN}      (√) ${ANSIColors.WHITE}Selected functions for execution: ${ANSIColors.MAGENTA}function1`
        )
      );
      chai.assert.isTrue(
        appendLineStub.calledWith(
          `${ANSIColors.GREEN}         (√) ${ANSIColors.WHITE}Function execution details: ${ANSIColors.GREEN}Status ${responseStatus}, ${ANSIColors.WHITE}refer to ${ANSIColors.BLUE}${logFilePath}${ANSIColors.WHITE} for all details.`
        )
      );
      chai.assert.isTrue(
        appendLineStub.calledWith(
          `${ANSIColors.RED}            (×) Error: ${ANSIColors.WHITE}Sample error`
        )
      );
    });
    it("should pretty print JSON", () => {
      const jsonText = '{"key":"value"}';
      const result = CopilotDebugLog.prettyPrintJson(jsonText);
      chai.assert.strictEqual(result, '{\n  "key": "value"\n}');
    });

    it("should parse full JSON with capabilities and initialize properties", () => {
      const logJson = JSON.stringify({
        pluginDeveloperInfo: {
          enabledPlugins: [{ name: "plugin1", id: "1", version: "1.0" }],
          matchedFunctionCandidates: [
            {
              plugin: { name: "plugin1", id: "1", version: "1.0" },
              functionDisplayName: "function1",
            },
          ],
          functionsSelectedForInvocation: [
            {
              plugin: { name: "plugin1", id: "1", version: "1.0" },
              functionDisplayName: "function1",
            },
          ],
          functionExecutions: [
            {
              function: {
                plugin: { name: "plugin1", id: "1", version: "1.0" },
                functionDisplayName: "function1",
              },
              executionStatus: { requestStatus: 200, responseStatus: 200, responseType: 1 },
              parameters: {},
              requestUri: "http://example.com",
              requestMethod: "GET",
              responseContent: "",
              responseContentType: "",
              errorMessage: "",
            },
          ],
        },
        capabilitiesDeveloperInfo: {
          enabledCapabilities: [
            {
              capabilityIcon: "iconUrl",
              capabilityName: "GraphConnectors",
              scopes: {},
            },
          ],
          capabilityExecutions: [
            {
              name: "GraphConnectors",
              status: "0",
              errorMessage: "",
            },
          ],
        },
        agentMetadata: {
          agentId: "agentId",
          agentVersion: "1.0",
          conversationId: "conversationId",
          requestId: "requestId",
        },
      });
      const copilotDebugLog = new CopilotDebugLog(logJson);
      chai.assert.deepEqual(copilotDebugLog.enabledPlugins, [
        { name: "plugin1", id: "1", version: "1.0" },
      ]);
      chai.assert.deepEqual(copilotDebugLog.matchedFunctionCandidates, [
        { plugin: { name: "plugin1", id: "1", version: "1.0" }, functionDisplayName: "function1" },
      ]);
      chai.assert.deepEqual(copilotDebugLog.functionsSelectedForInvocation, [
        { plugin: { name: "plugin1", id: "1", version: "1.0" }, functionDisplayName: "function1" },
      ]);
      chai.assert.deepEqual(copilotDebugLog.functionExecutions, [
        {
          function: {
            plugin: { name: "plugin1", id: "1", version: "1.0" },
            functionDisplayName: "function1",
          },
          executionStatus: { requestStatus: 200, responseStatus: 200, responseType: 1 },
          parameters: {},
          requestUri: "http://example.com",
          requestMethod: "GET",
          responseContent: "",
          responseContentType: "",
          errorMessage: "",
        },
      ]);
      chai.assert.deepEqual(copilotDebugLog.capabilitiesDeveloperInfo?.enabledCapabilities, [
        {
          capabilityIcon: "iconUrl",
          capabilityName: "GraphConnectors",
          scopes: {},
        },
      ]);
      chai.assert.deepEqual(copilotDebugLog.capabilitiesDeveloperInfo?.capabilityExecutions, [
        {
          name: "GraphConnectors",
          status: "0",
          errorMessage: "",
        },
      ]);
      chai.assert.deepEqual(copilotDebugLog.agentMetadata, {
        agentId: "agentId",
        agentVersion: "1.0",
        conversationId: "conversationId",
        requestId: "requestId",
      });
    });

    it("should write header details", () => {
      const logJson = JSON.stringify({
        agentMetadata: {
          agentId: "agentId",
          agentVersion: "1.0",
          conversationId: "conversationId",
          requestId: "requestId",
        },
      });
      const copilotDebugLog = new CopilotDebugLog(logJson);
      const appendLineStub = sandbox.stub(vscode.debug.activeDebugConsole, "appendLine");
      copilotDebugLog.write();
      chai.assert.isTrue(
        appendLineStub.calledWith(
          `${ANSIColors.WHITE}Agent ID (agentId). Conversation ID (conversationId). Request ID (requestId)`
        )
      );
      chai.assert.isTrue(
        appendLineStub.calledWith(
          `${
            ANSIColors.GREEN
          }${0} enabled capabilities, ${0} enabled actions, ${0} failed function executions, ${0} successful function executions, ${0} matched function candidates, ${0} functions selected for invocation.`
        )
      );
      chai.assert.isTrue(appendLineStub.calledWith(ANSIColors.WHITE + "Execution summary"));
    });

    it("should write with 0 capabilities", () => {
      const logJson = JSON.stringify({
        capabilitiesDeveloperInfo: {
          enabledCapabilities: [],
          capabilitiesExecutions: [],
        },
      });

      const copilotDebugLog = new CopilotDebugLog(logJson);
      const appendLineStub = sandbox.stub(vscode.debug.activeDebugConsole, "appendLine");
      copilotDebugLog.write();

      chai.assert.isTrue(appendLineStub.calledWith(ANSIColors.WHITE + "CAPABILITIES"));
      chai.assert.isTrue(
        appendLineStub.calledWith(
          `${ANSIColors.RED} (×) ${ANSIColors.WHITE}Enabled capabilities: None.`
        )
      );
      chai.assert.isTrue(
        appendLineStub.calledWith(
          `   ${ANSIColors.RED} (×) Execution status: ${ANSIColors.WHITE}None.`
        )
      );
    });

    it("should write with capabilities", () => {
      const logJson = JSON.stringify({
        capabilitiesDeveloperInfo: {
          enabledCapabilities: [
            {
              capabilityIcon: "iconUrl",
              capabilityName: "GraphConnectors",
              scopes: {
                scope1: {
                  scopeName: "scope1",
                },
              },
            },
          ],
          capabilitiesExecutions: [],
        },
      });

      const copilotDebugLog = new CopilotDebugLog(logJson);
      const appendLineStub = sandbox.stub(vscode.debug.activeDebugConsole, "appendLine");
      copilotDebugLog.write();

      chai.assert.isTrue(
        appendLineStub.calledWith(
          `${ANSIColors.GREEN}(√) ${ANSIColors.WHITE}Enabled capabilities: ${ANSIColors.MAGENTA}GraphConnectors`
        )
      );
    });

    it("should write with scoped capabilities", () => {
      const logJson = JSON.stringify({
        capabilitiesDeveloperInfo: {
          enabledCapabilities: [
            {
              capabilityIcon: "iconUrl",
              capabilityName: "GraphConnectors",
              scopes: {
                scope1: {
                  scopeName: "scope1",
                },
              },
            },
          ],
          capabilitiesExecutions: [],
        },
      });

      const copilotDebugLog = new CopilotDebugLog(logJson);
      const appendLineStub = sandbox.stub(vscode.debug.activeDebugConsole, "appendLine");
      copilotDebugLog.write();

      chai.assert.isTrue(
        appendLineStub.calledWith(`       ${ANSIColors.WHITE}scope1 - {"scopeName":"scope1"}`)
      );
    });

    it("should write with capabilities executions", () => {
      const logJson = JSON.stringify({
        capabilitiesDeveloperInfo: {
          enabledCapabilities: [
            {
              capabilityIcon: "iconUrl",
              capabilityName: "GraphConnectors",
              scopes: {
                scope1: {
                  scopeName: "scope1",
                },
              },
            },
          ],
          capabilitiesExecutions: [
            {
              name: "GraphConnectors",
              status: "1",
              errorMessage: "",
            },
          ],
        },
      });

      const fs = require("fs");
      sandbox.stub(globalVariables, "defaultExtensionLogPath").value("/path/to/log");
      sandbox.stub(Date.prototype, "toISOString").returns("test");
      const copilotDebugLog = new CopilotDebugLog(logJson);
      const appendLineStub = sandbox.stub(vscode.debug.activeDebugConsole, "appendLine");
      copilotDebugLog.write();

      chai.assert.isTrue(
        appendLineStub.calledWith(
          `${ANSIColors.GREEN}(√) ${ANSIColors.WHITE}Enabled capabilities: ${ANSIColors.MAGENTA}GraphConnectors`
        )
      );
    });
  });
});
