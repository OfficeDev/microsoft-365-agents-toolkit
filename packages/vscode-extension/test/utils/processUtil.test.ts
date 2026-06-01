import sinon, { SinonFakeTimers, useFakeTimers } from "sinon";
import * as chai from "chai";
import { execModule, killModule, processUtil, timeoutPromise } from "../../src/utils/processUtil";
describe("ProcessUtil", () => {
  const sandbox = sinon.createSandbox();

  afterEach(() => {
    sandbox.restore();
  });

  describe("killProcess", () => {
    it("error", async () => {
      const killStub = sandbox.stub(killModule, "killTree");
      killStub.yields(new Error());
      try {
        await processUtil.killProcess(-1, 5000, false);
        chai.assert.fail("Expected promise to reject, but it resolved.");
      } catch (error) {
        chai.assert.isTrue(error instanceof Error);
      }
    });
    it("happy", async () => {
      const killStub = sandbox.stub(killModule, "killTree");
      killStub.yields(null);
      await processUtil.killProcess(-1);
      chai.assert.isTrue(killStub.calledOnce);
    });
  });

  describe("getProcessIdsByPort", () => {
    it("should return PIDs from netstat output on Windows", async () => {
      const execStub = sandbox.stub(execModule, "exec") as sinon.SinonStub;
      const osStub = sandbox.stub(require("os"), "platform").returns("win32");
      execStub.callsFake((_cmd: string, _opts: any, cb: (...args: unknown[]) => void) => {
        cb(null, "  TCP    0.0.0.0:3978    0.0.0.0:0    LISTENING    12345\n");
      });
      const pids = await processUtil.getProcessIdsByPort(3978);
      chai.assert.deepEqual(pids, [12345]);
      osStub.restore();
    });

    it("should not match similar port numbers on Windows", async () => {
      const execStub = sandbox.stub(execModule, "exec") as sinon.SinonStub;
      const osStub = sandbox.stub(require("os"), "platform").returns("win32");
      execStub.callsFake((_cmd: string, _opts: any, cb: (...args: unknown[]) => void) => {
        cb(
          null,
          "  TCP    0.0.0.0:39780    0.0.0.0:0    LISTENING    99999\n  TCP    0.0.0.0:3978    0.0.0.0:0    LISTENING    12345\n"
        );
      });
      const pids = await processUtil.getProcessIdsByPort(3978);
      chai.assert.deepEqual(pids, [12345]);
      osStub.restore();
    });

    it("should return PIDs from lsof output on macOS", async () => {
      const execStub = sandbox.stub(execModule, "exec") as sinon.SinonStub;
      const osStub = sandbox.stub(require("os"), "platform").returns("darwin");
      execStub.callsFake((_cmd: string, _opts: any, cb: (...args: unknown[]) => void) => {
        cb(null, "12345\n67890\n");
      });
      const pids = await processUtil.getProcessIdsByPort(3978);
      chai.assert.deepEqual(pids, [12345, 67890]);
      osStub.restore();
    });

    it("should parse ss output on Linux when lsof is unavailable", async () => {
      const execStub = sandbox.stub(execModule, "exec") as sinon.SinonStub;
      const osStub = sandbox.stub(require("os"), "platform").returns("linux");
      execStub.callsFake((_cmd: string, _opts: any, cb: (...args: unknown[]) => void) => {
        cb(null, 'LISTEN  0  128  0.0.0.0:3978  0.0.0.0:*  users:(("node",pid=12345,fd=18))\n');
      });
      const pids = await processUtil.getProcessIdsByPort(3978);
      chai.assert.deepEqual(pids, [12345]);
      osStub.restore();
    });

    it("should return empty array on error", async () => {
      const execStub = sandbox.stub(execModule, "exec") as sinon.SinonStub;
      execStub.callsFake((_cmd: string, _opts: any, cb: (...args: unknown[]) => void) => {
        cb(new Error("command failed"), "");
      });
      const pids = await processUtil.getProcessIdsByPort(3978);
      chai.assert.deepEqual(pids, []);
    });

    it("should deduplicate PIDs on Windows", async () => {
      const execStub = sandbox.stub(execModule, "exec") as sinon.SinonStub;
      const osStub = sandbox.stub(require("os"), "platform").returns("win32");
      execStub.callsFake((_cmd: string, _opts: any, cb: (...args: unknown[]) => void) => {
        cb(
          null,
          "  TCP    0.0.0.0:3978    0.0.0.0:0    LISTENING    12345\n  TCP    [::]:3978    [::]:0    LISTENING    12345\n"
        );
      });
      const pids = await processUtil.getProcessIdsByPort(3978);
      chai.assert.deepEqual(pids, [12345]);
      osStub.restore();
    });
  });
});

describe("timeoutPromise", () => {
  let clock: SinonFakeTimers;

  beforeEach(() => {
    clock = useFakeTimers();
  });

  afterEach(() => {
    clock.restore();
  });

  it("timeoutPromise", async () => {
    try {
      const timeout = 1000;
      const promise = timeoutPromise(timeout, false);
      clock.tick(timeout);
      await promise;
      chai.assert.fail("Expected promise to reject, but it resolved.");
    } catch (error) {
      chai.assert.isTrue(error instanceof Error);
      chai.assert.equal(error.message, "Operation timeout");
    }
  });
  it("timeoutPromise - silent", async () => {
    try {
      const timeout = 1000;
      const promise = timeoutPromise(timeout, true);
      clock.tick(timeout);
      await promise;
    } catch (error) {
      chai.assert.fail("Expected promise to resolve, but it rejected.");
    }
  });
});
