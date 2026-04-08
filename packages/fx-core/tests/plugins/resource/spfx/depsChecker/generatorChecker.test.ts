// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { LogLevel, LogProvider, UserError } from "@microsoft/teamsfx-api";
import chai from "chai";
import fs from "fs-extra";
import "mocha";
import { restore, stub } from "sinon";
import { createContext, setTools } from "../../../../../src/common/globalVars";
import { cpUtils } from "../../../../../src/component/deps-checker/util/cpUtils";
import { GeneratorChecker } from "../../../../../src/component/generator/spfx/depsChecker/generatorChecker";
import { telemetryHelper } from "../../../../../src/component/generator/spfx/utils/telemetry-helper";
import { MockTools } from "../../../../core/utils";

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

describe("generator checker", () => {
  setTools(new MockTools());
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

  describe("getDependencyInfo", async () => {
    it("install", async () => {
      const generatorChecker = new GeneratorChecker(new StubLogger());
      const cleanStub = stub(GeneratorChecker.prototype, "cleanup" as any).callsFake(async () => {
        console.log("stub cleanup");
        return;
      });
      stub(cpUtils, "executeCommand").resolves();
      stub(fs, "pathExists").callsFake(async () => {
        return true;
      });

      try {
        await generatorChecker.install("1.18.2");
      } catch {
        chai.expect(cleanStub.callCount).equal(2);
      }
    });

    it("install error", async () => {
      const generatorChecker = new GeneratorChecker(new StubLogger());
      stub(GeneratorChecker.prototype, "cleanup" as any).callsFake(async () => {
        console.log("stub cleanup");
        return;
      });
      stub(cpUtils, "executeCommand").throws(new Error("unknown"));
      stub(fs, "pathExists").callsFake(async () => {
        return true;
      });

      let hasError = false;

      try {
        await generatorChecker.install("1.18.2");
      } catch {
        hasError = true;
      }

      chai.expect(hasError).to.be.true;
    });

    it("findGloballyInstalledVersion: returns version", async () => {
      const generatorChecker = new GeneratorChecker(new StubLogger());
      stub(cpUtils, "executeCommand").resolves(
        "C:\\Roaming\\npm\n`-- @microsoft/generator-sharepoint@1.16.1\n\n"
      );

      const res = await generatorChecker.findGloballyInstalledVersion(1);
      chai.expect(res).equal("1.16.1");
    });

    it("findGloballyInstalledVersion: regex error", async () => {
      const generatorChecker = new GeneratorChecker(new StubLogger());
      stub(cpUtils, "executeCommand").resolves(
        "C:\\Roaming\\npm\n`-- @microsoft/generator-sharepoint@empty\n\n"
      );

      const res = await generatorChecker.findGloballyInstalledVersion(1);
      chai.expect(res).equal(undefined);
    });

    it("findGloballyInstalledVersion: exeute commmand error", async () => {
      const generatorChecker = new GeneratorChecker(new StubLogger());
      stub(cpUtils, "executeCommand").throws("run command error");
      let error = undefined;

      try {
        const res = await generatorChecker.findGloballyInstalledVersion(1);
      } catch (e) {
        error = e;
      }
      chai.expect(error).not.undefined;
    });

    it("findLatestVersion: returns version", async () => {
      const generatorChecker = new GeneratorChecker(new StubLogger());
      stub(cpUtils, "executeCommand").resolves("1.16.1");

      const res = await generatorChecker.findLatestVersion(1);
      chai.expect(res).equal("1.16.1");
    });

    it("findLatestVersion: regex error", async () => {
      const generatorChecker = new GeneratorChecker(new StubLogger());
      stub(cpUtils, "executeCommand").resolves("empty");

      const res = await generatorChecker.findLatestVersion(1);
      chai.expect(res).to.be.undefined;
    });

    it("findLatestVersion: exeute commmand error", async () => {
      const generatorChecker = new GeneratorChecker(new StubLogger());
      stub(cpUtils, "executeCommand").throws("run command error");

      const res = await generatorChecker.findLatestVersion();
      chai.expect(res).to.be.undefined;
    });
  });

  describe("isLatestInstalled", () => {
    it("is latest installed", async () => {
      const checker = new GeneratorChecker(new StubLogger());
      stub(fs, "pathExists").callsFake(async () => {
        console.log("stub pathExists");
        return true;
      });

      stub(GeneratorChecker.prototype, "queryVersion" as any).callsFake(async () => {
        console.log("stub queryversion");
        return "latest";
      });

      stub(GeneratorChecker.prototype, "findLatestVersion" as any).callsFake(async () => {
        console.log("stub findLatestVersion");
        return "latest";
      });

      const result = await checker.isLatestInstalled("latest");
      chai.expect(result).is.true;
    });

    it("miss sentinel version", async () => {
      const checker = new GeneratorChecker(new StubLogger());
      stub(fs, "pathExists").callsFake(async () => {
        return false;
      });

      stub(GeneratorChecker.prototype, "queryVersion" as any).callsFake(async () => {
        return "latest";
      });

      stub(GeneratorChecker.prototype, "findLatestVersion" as any).callsFake(async () => {
        return "latest";
      });

      const result = await checker.isLatestInstalled("latest");
      chai.expect(result).is.false;
    });

    it("latest not installed", async () => {
      const checker = new GeneratorChecker(new StubLogger());
      stub(fs, "pathExists").callsFake(async () => {
        console.log("stub pathExists");
        return true;
      });

      stub(GeneratorChecker.prototype, "queryVersion" as any).callsFake(async () => {
        console.log("stub queryversion");
        return "lower version";
      });

      stub(GeneratorChecker.prototype, "findLatestVersion" as any).callsFake(async () => {
        console.log("stub findLatestVersion");
        return "latest";
      });

      const result = await checker.isLatestInstalled(undefined);
      chai.expect(result).is.false;
    });

    it("throw error", async () => {
      const checker = new GeneratorChecker(new StubLogger());
      stub(fs, "pathExists").callsFake(async () => {
        console.log("stub pathExists");
        return true;
      });

      stub(GeneratorChecker.prototype, "queryVersion" as any).throws("error");
      stub(GeneratorChecker.prototype, "findLatestVersion" as any).throws("error");

      const result = await checker.isLatestInstalled(undefined);
      chai.expect(result).is.false;
    });
  });

  describe("ensureDependency", () => {
    it("install successfully", async () => {
      const checker = new GeneratorChecker(new StubLogger());

      stub(GeneratorChecker.prototype, "install" as any).callsFake(async () => {
        console.log("installing");
      });

      const context = createContext();

      const result = await checker.ensureDependency(context, "1.18.2");
      chai.expect(result.isOk()).to.be.true;
    });

    it("install error", async () => {
      const checker = new GeneratorChecker(new StubLogger());

      stub(GeneratorChecker.prototype, "install" as any).callsFake(async () => {
        throw new UserError("source", "name", "msg", "msg");
      });

      const context = createContext();

      const result = await checker.ensureDependency(context, "1.18.2");
      chai.expect(result.isErr()).to.be.true;
    });
  });
});
