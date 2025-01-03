import * as chai from "chai";
import * as sinon from "sinon";
import * as vscode from "vscode";
import { CopilotDebugLog } from "../../src/pluginDebugger/copilotDebugLogOutput";
import { WebSocketEventHandler } from "../../src/pluginDebugger/webSocketEventHandler";
import * as ui from "../../src/qm/vsc_ui";

describe("WebSocketEventHandler", () => {
  const sandbox = sinon.createSandbox();

  beforeEach(() => {});

  afterEach(() => {
    sandbox.restore();
  });

  describe("handleEvent", () => {
    it("isWebSocketDataRelevant returns false", () => {
      sandbox.stub(WebSocketEventHandler, "isWebSocketDataRelevant").returns(false);
      const num = WebSocketEventHandler.handleEvent({ payloadData: '{"type":1' } as any);
      chai.assert.equal(num, 0);
    });
    it("throw error", () => {
      const appendLineStub = sandbox.stub(vscode.debug.activeDebugConsole, "appendLine");
      const mockUi = { showMessage: () => {} } as any;
      sandbox.stub(ui, "VS_CODE_UI").value(mockUi);
      const showMessageStub = sandbox.stub(mockUi, "showMessage");
      sandbox.stub(WebSocketEventHandler, "isWebSocketDataRelevant").returns(true);
      sandbox.stub(WebSocketEventHandler, "splitObjects").throws(new Error("Test"));
      const num = WebSocketEventHandler.handleEvent({ payloadData: '{"type":1' } as any);
      chai.assert.equal(num, 0);
      chai.assert.isTrue(showMessageStub.calledOnce);
      chai.assert.isTrue(appendLineStub.calledOnce);
    });
    it("happy", () => {
      sandbox.stub(WebSocketEventHandler, "isWebSocketDataRelevant").returns(true);
      const obj = { item: { messages: [] } };
      sandbox.stub(WebSocketEventHandler, "splitObjects").returns([JSON.stringify(obj)]);
      sandbox.stub(WebSocketEventHandler, "selectBotTextMessages").returns([{} as any]);
      sandbox.stub(WebSocketEventHandler, "convertBotMessageToChannelOutput").returns();
      const num = WebSocketEventHandler.handleEvent({ payloadData: '{"type":1' } as any);
      chai.assert.equal(num, 1);
    });
  });
  describe("isWebSocketDataRelevant", () => {
    it("true", () => {
      const res = WebSocketEventHandler.isWebSocketDataRelevant({
        payloadData: '{"type":2',
      } as any);
      chai.assert.isTrue(res);
    });
    it("false", () => {
      const res = WebSocketEventHandler.isWebSocketDataRelevant({
        payloadData: '{"type":1',
      } as any);
      chai.assert.isFalse(res);
    });
  });
  describe("splitObjects", () => {
    it("happy", () => {
      const res = WebSocketEventHandler.splitObjects({
        payloadData: "abc\x1e123",
      } as any);
      chai.assert.deepEqual(res, ["abc", "123"]);
    });
  });
  describe("selectBotTextMessages", () => {
    it("happy", () => {
      const res = WebSocketEventHandler.selectBotTextMessages({
        item: { messages: [{ messageType: "DeveloperLogs" }] },
      } as any);
      chai.assert.deepEqual(res, [{ messageType: "DeveloperLogs" }] as any);
    });
  });
  describe("convertBotMessageToChannelOutput", () => {
    it("happy", () => {
      const stub = sandbox.stub(CopilotDebugLog.prototype, "write");
      WebSocketEventHandler.convertBotMessageToChannelOutput({
        messageType: "DeveloperLogs",
        text: JSON.stringify({
          functionExecutions: [{ requestUrl: "" }],
        }),
      } as any);
      chai.assert.isTrue(stub.calledOnce);
    });
  });
  describe("convertBotMessageToChannelOutputJson", () => {
    it("happy", () => {
      const stub = sandbox.stub(WebSocketEventHandler, "prettyPrintJson");
      stub.returns(
        JSON.stringify({
          functionExecutions: [{ requestUrl: "" }],
        })
      );
      WebSocketEventHandler.convertBotMessageToChannelOutputJson({
        messageType: "DeveloperLogs",
        text: JSON.stringify({
          functionExecutions: [{ requestUrl: "" }],
        }),
      } as any);
      chai.assert.isTrue(stub.calledOnce);
    });
  });
  describe("prettyPrintJson", () => {
    it("happy", () => {
      const res = WebSocketEventHandler.prettyPrintJson(JSON.stringify({ a: "b" }));
      chai.assert.equal(res, JSON.stringify({ a: "b" }, null, 2));
    });
  });
});
