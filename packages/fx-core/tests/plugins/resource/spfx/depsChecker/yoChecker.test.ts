// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { LogLevel, LogProvider, UserError } from "@microsoft/teamsfx-api";
import { expect } from "chai";
import fs from "fs-extra";
import "mocha";
import rewire from "rewire";
import { assert, restore, spy, stub } from "sinon";
import { createContext, setTools } from "../../../../../src/common/globalVars";
import { cpUtils } from "../../../../../src/component/deps-checker/util/cpUtils";
import { YoChecker } from "../../../../../src/component/generator/spfx/depsChecker/yoChecker";
import { telemetryHelper } from "../../../../../src/component/generator/spfx/utils/telemetry-helper";
import { Utils } from "../../../../../src/component/generator/spfx/utils/utils";
import { MockTools } from "../../../../core/utils";

const ryc = rewire("../../../../../src/component/generator/spfx/depsChecker/yoChecker");

class StubLogger implements LogProvider {
  msg = "";
  verbose(msg: string): void {
    this.log(LogLevel.Verbose, msg);
  }
  debug(msg: string): void {
    this.log(LogLevel.Debug, msg);
  }
  info(msg: string | Array<any>): void {
    this.log(LogLevel.Info, msg as string);
  }
  warning(msg: string): void {
    this.log(LogLevel.Warning, msg);
  }
  error(msg: string): void {
    this.log(LogLevel.Error, msg);
  }
  log(level: LogLevel, msg: string): void {
    this.msg = msg;
  }
  async logInFile(level: LogLevel, msg: string): Promise<void> {
    this.msg = msg;
  }
  getLogFilePath(): string {
    return "";
  }
}

describe("Yo checker", () => {
  beforeEach(() => {
    stub(telemetryHelper, "sendSuccessEvent").callsFake(() => {
      console.log("success event");
      return;
    });
    stub(telemetryHelper, "sendErrorEvent").callsFake(() => {
      console.log("error event");
      return;
    });
  });

  afterEach(() => {
    restore();
  });

  it("install", async () => {
    const yc = new YoChecker(new StubLogger());
    const cleanStub = stub(YoChecker.prototype, "cleanup" as any).callsFake(async () => {
      console.log("stub cleanup");
      return;
    });
    stub(cpUtils, "executeCommand").resolves();
    stub(fs, "pathExists").callsFake(async () => {
      return true;
    });

    try {
      await yc.install("latest");
    } catch {
      assert.callCount(cleanStub, 2);
    }
  });

  it("install throw error", async () => {
    const yc = new YoChecker(new StubLogger());
    const cleanStub = stub(YoChecker.prototype, "cleanup" as any).callsFake(async () => {
      console.log("stub cleanup");
      return;
    });
    stub(cpUtils, "executeCommand").throws(new Error("unknown"));
    stub(fs, "pathExists").callsFake(async () => {
      return true;
    });

    try {
      await yc.install("latest");
    } catch (e) {
      expect(e.name).equal("NpmInstallError");
    }
  });

  it("clean up failed when install", async () => {
    const yc = new YoChecker(new StubLogger());
    stub(fs, "existsSync").returns(false);
    stub(fs, "emptyDir").throws("Failed to empty dir");
    stub(cpUtils, "executeCommand").resolves();
    stub(fs, "pathExists").callsFake(async () => {
      return true;
    });
    const logErrorSpy = spy(StubLogger.prototype, "error");

    await yc.install("latest");

    assert.callCount(logErrorSpy, 1);
  });

  it("findGloballyInstalledVersion: returns version", async () => {
    const generatorChecker = new YoChecker(new StubLogger());
    stub(cpUtils, "executeCommand").resolves("C:\\Roaming\\npm\n`-- yo@4.3.1\n\n");

    const res = await generatorChecker.findGloballyInstalledVersion(1);
    expect(res).equal("4.3.1");
  });

  it("findGloballyInstalledVersion: regex error", async () => {
    const yoChecker = new YoChecker(new StubLogger());
    stub(cpUtils, "executeCommand").resolves(
      "C:\\Roaming\\npm\n`-- @microsoft/generator-sharepoint@1.16.1\n\n"
    );

    const res = await yoChecker.findGloballyInstalledVersion(1);
    expect(res).equal(undefined);
  });

  it("findLatestVersion: returns version", async () => {
    const yoChecker = new YoChecker(new StubLogger());
    stub(cpUtils, "executeCommand").resolves("4.3.1");

    const res = await yoChecker.findLatestVersion(1);
    expect(res).equal("4.3.1");
  });

  it("findLatestVersion: regex error", async () => {
    const yoChecker = new YoChecker(new StubLogger());
    stub(cpUtils, "executeCommand").resolves("empty");

    const res = await yoChecker.findLatestVersion(1);
    expect(res).to.be.undefined;
  });

  it("findLatestVersion: exeute commmand error", async () => {
    const yoChecker = new YoChecker(new StubLogger());
    stub(cpUtils, "executeCommand").throws("run command error");

    const res = await yoChecker.findLatestVersion(1);
    expect(res).to.be.undefined;
  });

  describe("isLatestInstalled", () => {
    it("is latest installed", async () => {
      const yc = new YoChecker(new StubLogger());
      stub(fs, "pathExists").callsFake(async () => {
        console.log("stub pathExists");
        return true;
      });

      stub(YoChecker.prototype, "queryVersion" as any).callsFake(async () => {
        console.log("stub queryversion");
        return "latest";
      });

      stub(YoChecker.prototype, "findLatestVersion" as any).callsFake(async () => {
        console.log("stub findLatestVersion");
        return "latest";
      });

      const result = await yc.isLatestInstalled();
      expect(result).is.true;
    });

    it("latest not installed", async () => {
      const yc = new YoChecker(new StubLogger());
      stub(fs, "pathExists").callsFake(async () => {
        console.log("stub pathExists");
        return true;
      });

      stub(YoChecker.prototype, "queryVersion" as any).callsFake(async () => {
        console.log("stub queryversion");
        return "lowerVersion";
      });

      stub(YoChecker.prototype, "findLatestVersion" as any).callsFake(async () => {
        console.log("stub findLatestVersion");
        return "latest version";
      });

      const result = await yc.isLatestInstalled();
      expect(result).is.false;
    });

    it("sentitel file missing", async () => {
      const yc = new YoChecker(new StubLogger());
      stub(fs, "pathExists").callsFake(async () => {
        console.log("stub pathExists");
        return false;
      });

      stub(YoChecker.prototype, "queryVersion" as any).callsFake(async () => {
        console.log("stub queryversion");
        return "lowerVersion";
      });

      stub(YoChecker.prototype, "findLatestVersion" as any).callsFake(async () => {
        console.log("stub findLatestVersion");
        return "latest version";
      });

      const result = await yc.isLatestInstalled();
      expect(result).is.false;
    });

    it("throw error", async () => {
      const yc = new YoChecker(new StubLogger());
      stub(fs, "pathExists").callsFake(async () => {
        console.log("stub pathExists");
        return true;
      });

      stub(YoChecker.prototype, "queryVersion" as any).throws("error");
      stub(Utils, "findLatestVersion").throws("error");

      const result = await yc.isLatestInstalled();
      expect(result).is.false;
    });
  });

  describe("ensureDependency", () => {
    setTools(new MockTools());
    it("install successfully", async () => {
      const yc = new YoChecker(new StubLogger());

      stub(YoChecker.prototype, "install" as any).callsFake(async () => {
        console.log("installing");
      });

      const context = createContext();

      const result = await yc.ensureDependency(context, "latest");
      expect(result.isOk()).to.be.true;
    });

    it("install error", async () => {
      const yc = new YoChecker(new StubLogger());
      stub(YoChecker.prototype, "install" as any).callsFake(async () => {
        throw new UserError("source", "name", "msg", "msg");
      });

      const context = createContext();

      const result = await yc.ensureDependency(context, "latest");
      expect(result.isErr()).to.be.true;
    });
  });
});
