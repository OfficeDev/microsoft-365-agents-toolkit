import { assert } from "chai";
import * as fs from "fs-extra";
import path from "path";
import * as sinon from "sinon";
import { fileLockerDeps, withFileLock } from "../../../src/core/middleware/fileLocker";

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

    sandbox.stub(fileLockerDeps, "lock").callsFake(lockStub as any);
    sandbox.stub(fileLockerDeps, "waitSeconds").resolves();

    const result = await withFileLock(testFilePath, callback);

    assert.isTrue(callback.calledOnce);
    assert.strictEqual(result, "success");
    assert.strictEqual(lockStub.callCount, 2);
  });

  it("should throw an error if lock cannot be acquired after retries", async () => {
    const lockStub = sandbox.stub().throws({ code: "ELOCKED" });
    sandbox.stub(fileLockerDeps, "lock").callsFake(lockStub as any);
    sandbox.stub(fileLockerDeps, "waitSeconds").resolves();

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

  it("should throw an error if lock fails for a reason other than ELOCKED", async () => {
    const lockStub = sandbox.stub().throws(new Error("Some other error"));
    sandbox.stub(fileLockerDeps, "lock").callsFake(lockStub as any);
    try {
      await withFileLock(testFilePath, async () => "should not reach here");
      assert.fail("Expected error was not thrown");
    } catch (error) {
      assert.strictEqual(error.message, "Some other error");
    }
  });

  it("should release the lock after the callback is executed", async () => {
    const releaseStub = sandbox.stub().resolves();
    const lockStub = sandbox.stub().resolves(releaseStub);
    sandbox.stub(fileLockerDeps, "lock").callsFake(lockStub as any);

    const callback = sandbox.stub().resolves("success");

    const result = await withFileLock(testFilePath, callback);

    assert.isTrue(callback.calledOnce);
    assert.strictEqual(result, "success");
    assert.isTrue(releaseStub.calledOnce);
  });
});
