// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { ConversationReference, Activity } from "@microsoft/agents-activity";
import { assert, use as chaiUse } from "chai";
import * as chaiPromises from "chai-as-promised";
import * as sinon from "sinon";
import * as fs from "fs";
import * as path from "path";
import { DefaultConversationReferenceStore } from "../../../../../src/conversation/storage";
import { NotificationMiddleware } from "../../../../../src/conversation/middlewares/notificationMiddleware";

chaiUse(chaiPromises);

describe("Notification Middleware Tests - Node", () => {
  const sandbox = sinon.createSandbox();
  const fileDir = "./test/";
  const localFileName = ".notification.localstore.json";
  const filePath = path.join(fileDir, localFileName);
  const testData = {
    _a_1: {
      channelId: "1",
      conversation: {
        id: "1",
        tenantId: "a",
      },
    },
  };
  let middleware: NotificationMiddleware;

  beforeEach(() => {
    const conversationReferenceStore = new DefaultConversationReferenceStore(fileDir);
    middleware = new NotificationMiddleware({
      conversationReferenceStore,
    });

    sandbox.stub(Activity.prototype, "getConversationReference").callsFake(() => {
      const reference = {
        channelId: Activity.prototype.channelId,
        conversation: {
          id: Activity.prototype.conversation?.id,
          tenantId: Activity.prototype.conversation?.tenantId,
        },
      } as ConversationReference;
      if (Activity.prototype.conversation?.conversationType !== undefined) {
        reference.conversation.conversationType = Activity.prototype.conversation?.conversationType;
      }
      return reference;
    });
  });

  afterEach(() => {
    if (fs.existsSync(filePath)) {
      fs.unlinkSync(filePath);
    }
    sandbox.restore();
  });

  it("onTurn should correctly handle bot installed", async () => {
    const testContext = {
      activity: {
        action: "add",
        type: "installationUpdate",
        channelId: "1",
        conversation: {
          id: "1",
          tenantId: "a",
        },
        recipient: {
          id: "A",
        },
      },
    };
    await middleware.onTurn(testContext as any, async () => {});

    const fileData = JSON.parse(fs.readFileSync(filePath, "utf8"));
    assert.deepStrictEqual(fileData, {
      _a_1: {
        channelId: "1",
        conversation: {
          id: "1",
          tenantId: "a",
        },
      },
    });
  });

  it("onTurn should correctly handle bot uninstalled", async () => {
    fs.writeFileSync(filePath, JSON.stringify(testData), "utf8");
    const testContext = {
      activity: {
        action: "remove",
        type: "installationUpdate",
        channelId: "1",
        conversation: {
          id: "1",
          tenantId: "a",
        },
        recipient: {
          id: "A",
        },
      },
    };
    await middleware.onTurn(testContext as any, async () => {});
    const fileData = JSON.parse(fs.readFileSync(filePath, "utf8"));
    assert.deepStrictEqual(fileData, {});
  });

  it("onTurn should correctly handle bot messaged in general channel (new)", async () => {
    const testContext = {
      activity: {
        type: "message",
        channelId: "1",
        channelData: {
          team: {
            id: "X",
          },
        },
        conversation: {
          id: "1",
          conversationType: "channel",
          tenantId: "a",
        },
        recipient: {
          id: "A",
        },
      },
    };
    await middleware.onTurn(testContext as any, async () => {});
    const fileData = JSON.parse(fs.readFileSync(filePath, "utf8"));
    assert.deepStrictEqual(fileData, {
      _a_X: {
        channelId: "1",
        conversation: {
          id: "X",
          conversationType: "channel",
          tenantId: "a",
        },
      },
    });
  });

  it("onTurn should correctly handle bot messaged in general channel (exist)", async () => {
    const rawData = {
      _a_X: {
        channelId: "1",
        conversation: {
          id: "X",
          conversationType: "channel",
          tenantId: "a",
        },
      },
    };
    fs.writeFileSync(filePath, JSON.stringify(rawData), "utf8");

    const testContext = {
      activity: {
        type: "message",
        channelId: "xxxxxxxxx",
        channelData: {
          team: {
            id: "X",
          },
        },
        conversation: {
          id: "1",
          conversationType: "channel",
          tenantId: "a",
        },
        recipient: {
          id: "A",
        },
      },
    };
    await middleware.onTurn(testContext as any, async () => {});
    const fileData = JSON.parse(fs.readFileSync(filePath, "utf8"));
    assert.deepStrictEqual(fileData, rawData);
  });

  it("onTurn should correctly handle bot messaged in general channel (channelId)", async () => {
    const testContext = {
      activity: {
        type: "message",
        channelId: "1",
        channelData: {
          team: {
            id: "X",
          },
          channel: {
            id: "X",
          },
        },
        conversation: {
          id: "1",
          conversationType: "channel",
          tenantId: "a",
        },
        recipient: {
          id: "A",
        },
      },
    };
    await middleware.onTurn(testContext as any, async () => {});
    const fileData = JSON.parse(fs.readFileSync(filePath, "utf8"));
    assert.deepStrictEqual(fileData, {
      _a_X: {
        channelId: "1",
        conversation: {
          id: "X",
          conversationType: "channel",
          tenantId: "a",
        },
      },
    });
  });

  it("onTurn should ignore bot messaged in non-general channel", async () => {
    const testContext = {
      activity: {
        type: "message",
        channelId: "1",
        channelData: {
          team: {
            id: "X",
          },
          channel: {
            id: "X-channel",
          },
        },
        conversation: {
          id: "1",
          conversationType: "channel",
          tenantId: "a",
        },
        recipient: {
          id: "A",
        },
      },
    };
    await middleware.onTurn(testContext as any, async () => {});
    const fileData = fs.existsSync(filePath) ? JSON.parse(fs.readFileSync(filePath, "utf8")) : {};

    assert.deepStrictEqual(fileData, {});
  });

  it("onTurn should ignore bot messaged in channel (invalid data)", async () => {
    const testContext = {
      activity: {
        type: "message",
        channelId: "1",
        conversation: {
          id: "1",
          conversationType: "channel",
          tenantId: "a",
        },
        recipient: {
          id: "A",
        },
      },
    };
    await middleware.onTurn(testContext as any, async () => {});
    const fileData = fs.existsSync(filePath) ? JSON.parse(fs.readFileSync(filePath, "utf8")) : {};

    assert.deepStrictEqual(fileData, {});
  });

  it("onTurn should correctly handle bot messaged in chat (new)", async () => {
    const testContext = {
      activity: {
        type: "message",
        channelId: "1",
        conversation: {
          id: "1",
          conversationType: "groupChat",
          tenantId: "a",
        },
        recipient: {
          id: "A",
        },
      },
    };
    await middleware.onTurn(testContext as any, async () => {});
    const fileData = JSON.parse(fs.readFileSync(filePath, "utf8"));
    assert.deepStrictEqual(fileData, {
      _a_1: {
        channelId: "1",
        conversation: {
          id: "1",
          conversationType: "groupChat",
          tenantId: "a",
        },
      },
    });
  });

  it("onTurn should ignore bot messaged in chat (exist)", async () => {
    fs.writeFileSync(filePath, JSON.stringify(testData), "utf8");
    const testContext = {
      activity: {
        type: "message",
        channelId: "xxxxxxxxx",
        conversation: {
          id: "1",
          conversationType: "groupChat",
          tenantId: "a",
        },
        recipient: {
          id: "A",
        },
      },
    };
    await middleware.onTurn(testContext as any, async () => {});
    const fileData = JSON.parse(fs.readFileSync(filePath, "utf8"));
    assert.deepStrictEqual(fileData, testData);
  });

  it("onTurn should correctly handle team deleted", async () => {
    fs.writeFileSync(filePath, JSON.stringify(testData), "utf8");
    const testContext = {
      activity: {
        type: "conversationUpdate",
        channelId: "1",
        conversation: {
          id: "1",
          tenantId: "a",
        },
        recipient: {
          id: "A",
        },
        channelData: {
          eventType: "teamDeleted",
        },
      },
    };
    await middleware.onTurn(testContext as any, async () => {});
    const fileData = JSON.parse(fs.readFileSync(filePath, "utf8"));
    assert.deepStrictEqual(fileData, {});
  });

  it("onTurn should correctly handle team restored", async () => {
    const testContext = {
      activity: {
        type: "conversationUpdate",
        channelId: "1",
        conversation: {
          id: "1",
          tenantId: "a",
        },
        recipient: {
          id: "A",
        },
        channelData: {
          eventType: "teamRestored",
        },
      },
    };
    await middleware.onTurn(testContext as any, async () => {});
    const fileData = JSON.parse(fs.readFileSync(filePath, "utf8"));
    assert.deepStrictEqual(fileData, {
      _a_1: {
        channelId: "1",
        conversation: {
          id: "1",
          tenantId: "a",
        },
      },
    });
  });

  it("onTurn should ignore non-bot event", async () => {
    const testContext = {
      activity: {
        channelId: "1",
        recipient: {
          id: "B",
        },
      },
    };
    await middleware.onTurn(testContext as any, async () => {});
    const fileData = fs.existsSync(filePath) ? JSON.parse(fs.readFileSync(filePath, "utf8")) : {};

    assert.deepStrictEqual(fileData, {});
  });
});
