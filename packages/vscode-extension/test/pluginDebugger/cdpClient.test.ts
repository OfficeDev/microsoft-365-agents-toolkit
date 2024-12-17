import * as chai from "chai";
import sinon, { SinonFakeTimers, useFakeTimers } from "sinon";
import {
  CDPClient,
  cdpClientManager,
  CDPModule,
  isCopilotChatUrl,
  isM365ChatUrl,
  isM365CopilotChatDebugConfiguration,
  isOfficeChatUrl,
} from "../../src/pluginDebugger/cdpClient";
import { WebSocketEventHandler } from "../../src/pluginDebugger/webSocketEventHandler";
import * as ui from "../../src/qm/vsc_ui";
import { ExtTelemetry } from "../../src/telemetry/extTelemetry";

describe("cdpClient", () => {
  const sandbox = sinon.createSandbox();
  let clock: SinonFakeTimers;

  beforeEach(() => {
    clock = useFakeTimers();
  });

  afterEach(() => {
    sandbox.restore();
    clock.restore();
  });

  describe("connectWithBackoff", () => {
    it("build fail", async () => {
      sandbox.stub(CDPModule, "build").rejects(new Error());
      const client = new CDPClient("url", 9222, "name");
      try {
        const p = client.connectWithBackoff(9222, "", 1, 1);
        clock.tick(1);
        await p;
        chai.assert.fail("should not reach here");
      } catch (e) {
        chai.assert.isDefined(e);
      }
    });
  });
  describe("subscribeToWebSocketEvents", () => {
    it("happy", async () => {
      const cdpClient = new CDPClient("url", 9222, "name");
      sandbox.stub(cdpClient, "url").value("xxx");
      const client = {
        Network: { enable: () => {}, webSocketFrameReceived: () => {} },
        Page: { enable: () => {} },
        Target: {
          getTargets: () => {
            return { targetInfos: [] };
          },
        },
      } as any;
      const webSocketFrameReceived = sandbox.stub(client.Network, "webSocketFrameReceived");
      await cdpClient.subscribeToWebSocketEvents(client);
      chai.assert.isTrue(webSocketFrameReceived.called);
    });
    it("connect to iframe target", async () => {
      const cdpClient = new CDPClient("url", 9222, "name");
      sandbox.stub(cdpClient, "launchTeamsChatListener").resolves();
      const client = {
        Network: { enable: () => {}, webSocketFrameReceived: () => {} },
        Page: { enable: () => {} },
        Target: {
          getTargets: () => {
            return { targetInfos: [] };
          },
        },
      } as any;
      sandbox.stub(cdpClient, "url").value("m365.cloud.microsoft/chat");
      const webSocketFrameReceived = sandbox.stub(client.Network, "webSocketFrameReceived");
      await cdpClient.subscribeToWebSocketEvents(client);
      chai.assert.isTrue(webSocketFrameReceived.notCalled);
    });
  });
  describe("start", () => {
    it("happy", async () => {
      const sendTelemetryEvent = sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
      const cdpClient = new CDPClient("url", 9222, "name");
      sandbox.stub(CDPModule, "build").resolves({
        Network: { enable: () => {}, webSocketFrameReceived: () => {} },
        Page: { enable: () => {} },
        Target: {
          getTargets: () => {
            return { targetInfos: [] };
          },
        },
      } as any);
      cdpClient.errors = [new Error()];
      sandbox.stub(cdpClient, "subscribeToWebSocketEvents").resolves();
      const startPromise = cdpClient.start();
      clock.tick(2000);
      await startPromise;
      chai.assert.isTrue(sendTelemetryEvent.called);
    });
    it("error", async () => {
      const sendTelemetryEvent = sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
      const sendTelemetryErrorEvent = sandbox.stub(ExtTelemetry, "sendTelemetryErrorEvent");
      const cdpClient = new CDPClient("url", 9222, "name");
      cdpClient.errors = [new Error()];
      sandbox.stub(CDPModule, "build").resolves({
        Network: { enable: () => {}, webSocketFrameReceived: () => {} },
        Page: { enable: () => {} },
        Target: {
          getTargets: () => {
            return { targetInfos: [] };
          },
        },
      } as any);
      sandbox.stub(cdpClient, "subscribeToWebSocketEvents").rejects(new Error());
      const startPromise = cdpClient.start();
      clock.tick(2000);
      await startPromise;
      chai.assert.isTrue(sendTelemetryEvent.called);
      chai.assert.isTrue(sendTelemetryErrorEvent.called);
    });
  });
  describe("stop", () => {
    it("happy", async () => {
      const sendTelemetryEvent = sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
      const cdpClient = new CDPClient("url", 9222, "name");
      cdpClient.errors = [new Error()];
      cdpClient.client = {
        close: () => {},
      } as any;
      await cdpClient.stop();
      chai.assert.isTrue(sendTelemetryEvent.called);
    });
  });
  describe("webSocketFrameReceivedHandler", () => {
    it("happy", async () => {
      const stub = sandbox.stub(WebSocketEventHandler, "handleEvent");
      stub.returns(1);
      const cdpClient = new CDPClient("url", 9222, "name");
      cdpClient.webSocketFrameReceivedHandler({} as any);
      chai.assert.isTrue(stub.called);
    });
  });

  describe("launchTeamsChatListener", () => {
    it("happy", async () => {
      const cdpClient = new CDPClient("url", 9222, "name");
      const stub = sandbox.stub(cdpClient, "connectToTargetIframe");
      const client = {} as any;
      stub.resolves(true);
      cdpClient.launchTeamsChatListener(client);
      chai.assert.isTrue(stub.calledOnce);
    });
    // it("error", async () => {
    //   const cdpClient = new CDPClient("url", 9222, "name");
    //   sandbox.stub(cdpClient, "connectToTargetIframe").rejects(new Error());
    //   const client = {} as any;
    //   await cdpClient.launchTeamsChatListener(client, 1, 1);
    //   chai.assert.isUndefined(cdpClient.client);
    // });
    it("reach max try", async () => {
      const cdpClient = new CDPClient("url", 9222, "name");
      sandbox.stub(cdpClient, "connectToTargetIframe").rejects(new Error());
      const client = {} as any;
      const p = cdpClient.launchTeamsChatListener(client, 2, 1);
      await clock.tickAsync(2000);
      await p;
      chai.assert.isUndefined(cdpClient.client);
    });
  });

  describe("connectToTargetIframe", () => {
    it("no targetInfo", async () => {
      const client = {
        Network: { enable: () => {}, webSocketFrameReceived: () => {} },
        Page: { enable: () => {} },
        Target: {
          getTargets: () => {
            return { targetInfos: [] };
          },
        },
      } as any;
      const cdpClient = new CDPClient("url", 9222, "name");
      const res = await cdpClient.connectToTargetIframe(client);
      chai.assert.isFalse(res);
    });
    it("no sessionClient", async () => {
      const client = {
        Network: { enable: () => {}, webSocketFrameReceived: () => {} },
        Page: { enable: () => {} },
        Target: {
          getTargets: () => {
            return {
              targetInfos: [
                {
                  type: "iframe",
                  url: "outlook.office.com/hosted/semanticoverview/Users",
                },
              ],
            };
          },
        },
      } as any;
      const cdpClient = new CDPClient("url", 9222, "name");
      sandbox.stub(cdpClient, "connectWithBackoff").resolves(undefined);
      const res = await cdpClient.connectToTargetIframe(client);
      chai.assert.isFalse(res);
    });
    it("happy path", async () => {
      const client = {
        Network: { enable: () => {}, webSocketFrameReceived: () => {} },
        Page: { enable: () => {} },
        Target: {
          getTargets: () => {
            return {
              targetInfos: [
                {
                  type: "iframe",
                  url: "outlook.office.com/hosted/semanticoverview/Users",
                },
              ],
            };
          },
        },
        close: () => {},
      } as any;
      const cdpClient = new CDPClient("url", 9222, "name");
      sandbox.stub(cdpClient, "connectWithBackoff").resolves(client);
      const res = await cdpClient.connectToTargetIframe(client);
      chai.assert.isTrue(res);
    });
  });
});

describe("isM365CopilotChatDebugConfiguration", () => {
  it("true", async () => {
    const config: any = {
      request: "launch",
      url: "https://www.office.com/chat?auth=2&developerMode=Basic",
      runtimeArgs: ["--remote-debugging-port=9222"],
    };
    const res = isM365CopilotChatDebugConfiguration(config);
    chai.assert.isDefined(res);
  });

  it("false - request", async () => {
    const config: any = {
      request: "abc",
    };
    const res = isM365CopilotChatDebugConfiguration(config);
    chai.assert.isUndefined(res);
  });
  it("false - url undefined", async () => {
    const config: any = {
      request: "launch",
    };
    const res = isM365CopilotChatDebugConfiguration(config);
    chai.assert.isUndefined(res);
  });

  it("false - url is not chat", async () => {
    const config: any = {
      request: "launch",
      url: "https://abc",
    };
    const res = isM365CopilotChatDebugConfiguration(config);
    chai.assert.isUndefined(res);
  });

  it("false - url param", async () => {
    const config: any = {
      request: "launch",
      url: "https://www.office.com/chat?auth=2",
    };
    const res = isM365CopilotChatDebugConfiguration(config);
    chai.assert.isUndefined(res);
  });

  it("false - runtimeArgs undefined", async () => {
    const config: any = {
      request: "launch",
      url: "https://www.office.com/chat?auth=2&developerMode=Basic",
    };
    const res = isM365CopilotChatDebugConfiguration(config);
    chai.assert.isUndefined(res);
  });

  it("false - runtimeArgs not contains port", async () => {
    const config: any = {
      request: "launch",
      url: "https://www.office.com/chat?auth=2&developerMode=Basic",
      runtimeArgs: [],
    };
    const res = isM365CopilotChatDebugConfiguration(config);
    chai.assert.isUndefined(res);
  });

  it("false - runtimeArgs contains invalid port", async () => {
    const config: any = {
      request: "launch",
      url: "https://www.office.com/chat?auth=2&developerMode=Basic",
      runtimeArgs: ["--remote-debugging-port=abc"],
    };
    const res = isM365CopilotChatDebugConfiguration(config);
    chai.assert.isUndefined(res);
  });
});

describe("isCopilotChatUrl", () => {
  it("true", async () => {
    const res = isCopilotChatUrl("https://www.office.com/chat?auth=2&developerMode=Basic");
    chai.assert.isTrue(res);
  });

  it("false", async () => {
    const res = isCopilotChatUrl("https://abc.com");
    chai.assert.isFalse(res);
  });
});

describe("isOfficeChatUrl", () => {
  it("true", async () => {
    const res = isOfficeChatUrl("https://www.office.com/chat?auth=2&developerMode=Basic");
    chai.assert.isTrue(res);
  });

  it("false", async () => {
    const res = isOfficeChatUrl("https://abc.com");
    chai.assert.isFalse(res);
  });
});

describe("isM365ChatUrl", () => {
  it("true", async () => {
    const res = isM365ChatUrl("https://m365.cloud.microsoft/chat");
    chai.assert.isTrue(res);
  });
  it("false", async () => {
    const res = isM365ChatUrl("https://abc.com");
    chai.assert.isFalse(res);
  });
});

describe("CDPClientManager", () => {
  const sandbox = sinon.createSandbox();

  afterEach(() => {
    sandbox.restore();
  });

  describe("start", () => {
    it("exist", async () => {
      sandbox.stub(ui, "VS_CODE_UI").value({ showMessage: () => {} } as any);
      sandbox.stub(CDPClient.prototype, "stop").resolves();
      sandbox.stub(CDPClient.prototype, "start").resolves();
      cdpClientManager.sessions.set(9222, new CDPClient("url", 9222, "name"));
      const client = cdpClientManager.start("https://m365.cloud.microsoft/chat", 9222, "name");
      chai.assert.isDefined(client);
    });
  });
  describe("stop", () => {
    it("happy", async () => {
      const client = new CDPClient("url", 9222, "name");
      cdpClientManager.sessions.set(9222, client);
      const stub = sandbox.stub(client, "stop").resolves();
      await cdpClientManager.stop(9222);
      chai.assert.isTrue(stub.called);
    });
  });
});
