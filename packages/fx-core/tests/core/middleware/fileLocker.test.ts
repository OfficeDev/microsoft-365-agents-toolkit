import * as fs from "fs-extra";
import * as sinon from "sinon";
import { assert } from "chai";
import path from "path";
import { withFileLock } from "../../../src/core/middleware/fileLocker";
import "mocha";

describe("withFileLock", () => {
  const sandbox = sinon.createSandbox();
  const testFilePath = path.join(__dirname, "test.lock");

  beforeEach(async () => {
    await fs.ensureFile(testFilePath);
  });

  afterEach(async () => {
    sandbox.restore();
    await fs.remove(testFilePath);
  });

  it("should execute the callback when lock is acquired", async () => {
    const callback = sandbox.stub().resolves("success");

    const result = await withFileLock(testFilePath, callback);

    assert.isTrue(callback.calledOnce);
    assert.strictEqual(result, "success");
  });

  it("should throw an error if the file does not exist", async () => {
    await fs.remove(testFilePath);

    try {
      await withFileLock(testFilePath, async () => "should not reach here");
      assert.fail("Expected error was not thrown");
    } catch (error) {
      assert.strictEqual(error.message, `File not found: ${testFilePath}`);
    }

    await fs.ensureFile(testFilePath);
  });

  it("should retry acquiring the lock if it is already locked", async () => {
    const callback = sandbox.stub().resolves("success");
    const lockStub = sandbox
      .stub()
      .onFirstCall()
      .throws({ code: "ELOCKED" })
      .onSecondCall()
      .resolves(async () => {});

    sandbox.replace(require("proper-lockfile"), "lock", lockStub);

    const result = await withFileLock(testFilePath, callback);

    assert.isTrue(callback.calledOnce);
    assert.strictEqual(result, "success");
    assert.strictEqual(lockStub.callCount, 2);
  });

  it("should throw an error if lock cannot be acquired after retries", async () => {
    const lockStub = sandbox.stub().throws({ code: "ELOCKED" });
    sandbox.replace(require("proper-lockfile"), "lock", lockStub);

    try {
      await withFileLock(testFilePath, async () => "should not reach here");
      assert.fail("Expected error was not thrown");
    } catch (error) {
      assert.strictEqual(
        error.message,
        `Failed to acquire lock on ${testFilePath} after 10 seconds.`
      );
    }
  });

  it("should release the lock after the callback is executed", async () => {
    const releaseStub = sandbox.stub().resolves();
    const lockStub = sandbox.stub().resolves(releaseStub);
    sandbox.replace(require("proper-lockfile"), "lock", lockStub);

    const callback = sandbox.stub().resolves("success");

    const result = await withFileLock(testFilePath, callback);

    assert.isTrue(callback.calledOnce);
    assert.strictEqual(result, "success");
    assert.isTrue(releaseStub.calledOnce);
  });
});
