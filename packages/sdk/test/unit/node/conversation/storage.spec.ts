// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { ConversationReference } from "@microsoft/agents-activity";
import { assert, use as chaiUse } from "chai";
import * as chaiPromises from "chai-as-promised";
import * as fs from "fs";
import * as sinon from "sinon";
import { DefaultConversationReferenceStore } from "../../../../src/conversation/storage";
import * as path from "path";

chaiUse(chaiPromises);

describe("DefaultConversationReferenceStore Tests - Node", () => {
  const testDir = "./test/";
  const localFileName = ".notification.localstore.json";
  const testPath = path.join(testDir, localFileName);
  let testStore: DefaultConversationReferenceStore;

  beforeEach(() => {
    testStore = new DefaultConversationReferenceStore(testDir);
  });

  afterEach(() => {
    if (fs.existsSync(testPath)) {
      fs.unlinkSync(testPath);
    }
    sinon.restore();
  });

  it("list should return correct data", async () => {
    fs.writeFileSync(
      testPath,
      JSON.stringify({ _a_1: { conversation: { id: "1", tenantId: "a" } } }),
      "utf8"
    );
    const { data } = await testStore.list();
    assert.deepStrictEqual(data, [
      {
        conversation: {
          id: "1",
          tenantId: "a",
        },
      } as ConversationReference,
    ]);
  });

  it("list should return empty data if storage is empty", async () => {
    const { data } = await testStore.list();
    assert.strictEqual(data.length, 0);
  });

  it("add should persist correct data", async () => {
    await testStore.add(
      "_a_1",
      {
        conversation: {
          id: "1",
          tenantId: "a",
        },
      } as ConversationReference,
      { overwrite: true }
    );
    const fileData = JSON.parse(fs.readFileSync(testPath).toString());
    assert.deepStrictEqual(fileData, {
      _a_1: {
        conversation: {
          id: "1",
          tenantId: "a",
        },
      },
    });
  });

  it("add with overwrite should update existing data", async () => {
    fs.writeFileSync(
      testPath,
      JSON.stringify({
        _a_1: {
          channelId: "1",
          conversation: {
            id: "1",
            tenantId: "a",
          },
        },
      }),
      "utf8"
    );
    const added = await testStore.add(
      "_a_1",
      {
        channelId: "2",
        conversation: {
          id: "1",
          tenantId: "a",
        },
      } as ConversationReference,
      { overwrite: true }
    );
    assert.isTrue(added);
    const fileData = JSON.parse(fs.readFileSync(testPath).toString());
    assert.deepStrictEqual(fileData, {
      _a_1: {
        channelId: "2",
        conversation: {
          id: "1",
          tenantId: "a",
        },
      },
    });
  });

  it("add without overwrite should skip updating existing data", async () => {
    fs.writeFileSync(
      testPath,
      JSON.stringify({
        _a_1: {
          channelId: "1",
          conversation: {
            id: "1",
            tenantId: "a",
          },
        },
      }),
      "utf8"
    );
    const added = await testStore.add(
      "_a_1",
      {
        channelId: "2",
        conversation: {
          id: "1",
          tenantId: "a",
        },
      } as ConversationReference,
      { overwrite: false }
    );
    assert.isFalse(added);

    const fileData = JSON.parse(fs.readFileSync(testPath).toString());
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

  it("delete should remove correct data", async () => {
    fs.writeFileSync(
      testPath,
      JSON.stringify({ _a_1: { conversation: { id: "1", tenantId: "a" } } }),
      "utf8"
    );

    const removed = await testStore.remove("_a_1", {
      conversation: {
        id: "1",
        tenantId: "a",
      },
    } as ConversationReference);
    assert.isTrue(removed);

    const fileData = JSON.parse(fs.readFileSync(testPath).toString());
    assert.deepStrictEqual(fileData, {});
  });

  it("delete non-existing data should return correct result", async () => {
    const removed = await testStore.remove("_a_1", {
      conversation: {
        id: "1",
        tenantId: "a",
      },
    } as ConversationReference);
    assert.isFalse(removed);

    const fileData = fs.existsSync(testPath) ? JSON.parse(fs.readFileSync(testPath, "utf8")) : {};
    assert.deepStrictEqual(fileData, {});
  });
});
