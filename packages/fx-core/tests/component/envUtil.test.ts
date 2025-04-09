import { hooks } from "@feathersjs/hooks/lib";
import {
  err,
  FxError,
  Inputs,
  ok,
  Platform,
  Result,
  Settings,
  UserError,
} from "@microsoft/teamsfx-api";
import { assert } from "chai";
import fs from "fs-extra";
import "mocha";
import mockedEnv, { RestoreFn } from "mocked-env";
import * as path from "path";
import * as sinon from "sinon";
import { MetadataV3, MetadataV3Abandoned, MetadataV4 } from "../../src/common/versionMetadata";
import { ProjectModel } from "../../src/component/configManager/interface";
import { EnvLoaderMW, EnvWriterMW } from "../../src/component/middleware/envMW";
import { DotenvOutput, dotenvUtil, envUtil } from "../../src/component/utils/envUtil";
import { pathUtils } from "../../src/component/utils/pathUtils";
import { settingsUtil } from "../../src/component/utils/settingsUtil";
import { LocalCrypto } from "../../src/core/crypto";
import { environmentManager } from "../../src/core/environment";
import { FxCore } from "../../src/core/FxCore";
import { globalVars, setTools, TOOLS } from "../../src/common/globalVars";
import { ContextInjectorMW } from "../../src/core/middleware/contextInjector";
import { CoreHookContext } from "../../src/core/types";
import {
  FileNotFoundError,
  MissingEnvironmentVariablesError,
  MissingRequiredFileError,
  NoEnvFilesError,
  UserCancelError,
} from "../../src/error/common";
import { MockTools } from "../core/utils";
import { parseSetOutputCommand } from "../../src/component/driver/script/scriptDriver";
import * as yaml from "yaml";

describe("envUtils", () => {
  const tools = new MockTools();
  setTools(tools);
  const sandbox = sinon.createSandbox();
  const cryptoProvider = new LocalCrypto("mockProjectId");
  const decrypted = "123";
  const mockSettings: Settings = {
    trackingId: "mockProjectId",
    version: "1",
  };
  let mockedEnvRestore: RestoreFn = () => {};
  afterEach(() => {
    sandbox.restore();
    if (mockedEnvRestore) {
      mockedEnvRestore();
    }
  });

  describe("pathUtils.getYmlFilePath", () => {
    afterEach(() => {
      sandbox.restore();
      mockedEnvRestore();
    });
    it("defined in env var of TEAMSFX_CONFIG_FILE_PATH", async () => {
      mockedEnvRestore = mockedEnv({ TEAMSFX_CONFIG_FILE_PATH: "abc.yml" });
      const res1 = pathUtils.getYmlFilePath(".");
      assert.equal(res1, path.join(".", "abc.yml"));
    });
    it("happy path v4 dev by TEAMSFX_ENV", async () => {
      mockedEnvRestore = mockedEnv({ TEAMSFX_ENV: "dev" });
      sandbox.stub(fs, "pathExistsSync").returns(true);
      const res1 = pathUtils.getYmlFilePath(".");
      assert.equal(res1, path.join(".", MetadataV4.configFile));
    });
    it("happy path v4 dev default", async () => {
      sandbox.stub(fs, "pathExistsSync").returns(true);
      const res1 = pathUtils.getYmlFilePath(".");
      assert.equal(res1, path.join(".", MetadataV4.configFile));
    });
    it("happy path v4 local", async () => {
      sandbox.stub(fs, "pathExistsSync").returns(true);
      const res1 = pathUtils.getYmlFilePath(".", "local");
      assert.equal(res1, path.join(".", MetadataV4.localConfigFile));
    });
    it("happy path v4 testtool", async () => {
      sandbox.stub(fs, "pathExistsSync").returns(true);
      const res1 = pathUtils.getYmlFilePath(".", "testtool");
      assert.equal(res1, path.join(".", MetadataV4.testToolConfigFile));
    });
    it("happy path v4 sandbox", async () => {
      sandbox.stub(fs, "pathExistsSync").returns(true);
      const res1 = pathUtils.getYmlFilePath(".", "sandbox");
      assert.equal(res1, path.join(".", MetadataV4.sandboxConfigFile));
    });
    it("happy path v3 local", async () => {
      sandbox.stub(fs, "pathExistsSync").onFirstCall().returns(false).onSecondCall().returns(true);
      const res1 = pathUtils.getYmlFilePath(".", "local");
      assert.equal(res1, path.join(".", MetadataV3.localConfigFile));
    });
    it("happy path v3 testtool", async () => {
      sandbox.stub(fs, "pathExistsSync").onFirstCall().returns(false).onSecondCall().returns(true);
      const res1 = pathUtils.getYmlFilePath(".", "testtool");
      assert.equal(res1, path.join(".", MetadataV3.testToolConfigFile));
    });
    it("happy path v3 sandbox", async () => {
      sandbox.stub(fs, "pathExistsSync").onFirstCall().returns(false).onSecondCall().returns(true);
      const res1 = pathUtils.getYmlFilePath(".", "sandbox");
      assert.equal(res1, path.join(".", MetadataV3.sandboxConfigFile));
    });
    it("throw MissingRequiredFileError with env=dev", async () => {
      sandbox.stub(fs, "pathExistsSync").returns(false);
      process.env.TEAMSFX_ENV = "dev";
      try {
        pathUtils.getYmlFilePath(".", "dev");
        assert.fail("show not reach here");
      } catch (e) {
        assert.isTrue(e instanceof MissingRequiredFileError);
      }
    });
    it("throw MissingRequiredFileError with env=local", async () => {
      sandbox.stub(fs, "pathExistsSync").returns(false);
      process.env.TEAMSFX_ENV = "local";
      try {
        await pathUtils.getYmlFilePath(".", "local");
        assert.fail("show not reach here");
      } catch (e) {
        assert.isTrue(e instanceof MissingRequiredFileError);
      }
    });
    it("happy path for customized yaml path", async () => {
      mockedEnvRestore = mockedEnv({ TEAMSFX_CONFIG_FILE_PATH: "./customized.yml" });
      const res1 = pathUtils.getYmlFilePath(".", "dev");
      assert.equal(res1, "./customized.yml");
    });
  });

  describe("pathUtils.getEnvFolderPath", () => {
    it("happy path", async () => {
      sandbox
        .stub(fs, "readFile")
        .resolves("version: 1.0.0\nenvironmentFolderPath: /home/envs" as any);
      sandbox.stub(fs, "pathExists").resolves(true);
      sandbox.stub(pathUtils, "getYmlFilePath").returns("./xxx");
      const res = await pathUtils.getEnvFolderPath(".");
      assert.isTrue(res.isOk());
      if (res.isOk()) {
        assert.equal(res.value, "/home/envs");
      }
    });
    it("returns default value", async () => {
      sandbox.stub(pathUtils, "getYmlFilePath").returns("./m365agents.yml");
      sandbox.stub(fs, "readFile").resolves("version: 1.0.0" as any);
      sandbox.stub(fs, "pathExists").resolves(true);
      const res = await pathUtils.getEnvFolderPath("");
      assert.isTrue(res.isOk());
    });
    it("returns undefined value", async () => {
      sandbox.stub(pathUtils, "getYmlFilePath").returns("./m365agents.yml");
      sandbox
        .stub(fs, "readFile")
        .resolves("version: 1.0.0\nenvironmentFolderPath: /home/envs" as any);
      sandbox.stub(fs, "pathExists").resolves(false);
      const res = await pathUtils.getEnvFolderPath("");
      assert.isTrue(res.isOk());
      if (res.isOk()) {
        assert.isUndefined(res.value);
      }
    });
  });

  describe("pathUtils.getEnvFilePath", () => {
    it("happy path", async () => {
      sandbox.stub(pathUtils, "getYmlFilePath").returns("./xxx");
      sandbox
        .stub(fs, "readFile")
        .resolves("version: 1.0.0\nenvironmentFolderPath: /home/envs" as any);
      sandbox.stub(fs, "pathExists").resolves(true);
      const res = await pathUtils.getEnvFilePath(".", "dev");
      assert.isTrue(res.isOk());
      if (res.isOk()) {
        assert.equal(res.value, path.join("/home/envs", ".env.dev"));
      }
    });
    it("returns default value", async () => {
      sandbox.stub(fs, "readFile").resolves("version: 1.0.0" as any);
      sandbox.stub(fs, "pathExists").resolves(true);
      sandbox.stub(pathUtils, "getYmlFilePath").returns("./xxx");
      const res = await pathUtils.getEnvFilePath(".", "dev");
      assert.isTrue(res.isOk());
    });
  });
  describe("pathUtils.resolveFilePath", () => {
    const sandbox = sinon.createSandbox();
    beforeEach(() => {});
    afterEach(async () => {
      sandbox.restore();
    });
    it("relative path", async () => {
      const res = pathUtils.resolveFilePath("/test");
      assert.isTrue(res === "\\test" || res === "/test");
    });
    it("absolute path", async () => {
      const res = pathUtils.resolveFilePath("/test", "/test2");
      assert.isTrue(res === "\\test2" || res === "/test2");
    });
  });
  describe("pathUtils.readEnv", () => {
    it("happy path", async () => {
      sandbox.stub(pathUtils, "getEnvFilePath").resolves(ok("."));
      const encRes = await cryptoProvider.encrypt(decrypted);
      if (encRes.isErr()) throw encRes.error;
      const encrypted = encRes.value;
      sandbox
        .stub(fs, "readFile")
        .onFirstCall()
        .resolves("TEAMSFX_ENV=env\nTEAMS_APP_ID=testappid\nTAB_ENDPOINT=testendpoint" as any)
        .onSecondCall()
        .resolves(("SECRET_ABC=" + encrypted) as any);
      sandbox.stub(fs, "pathExists").resolves(true);
      sandbox.stub(settingsUtil, "readSettings").resolves(ok(mockSettings));
      const res = await envUtil.readEnv(".", "dev");
      assert.isTrue(res.isOk());
      assert.equal(process.env.SECRET_ABC, decrypted);
    });
    it("silent", async () => {
      sandbox.stub(fs, "pathExists").resolves(false);
      sandbox.stub(pathUtils, "getEnvFilePath").resolves(ok("."));
      sandbox.stub(settingsUtil, "readSettings").resolves(ok(mockSettings));
      const res = await envUtil.readEnv(".", "dev", false, true);
      assert.isTrue(res.isOk());
      if (res.isOk()) {
        assert.deepEqual(res.value, { TEAMSFX_ENV: "dev" });
      }
    });
    it("not silent 1", async () => {
      sandbox.stub(fs, "pathExists").resolves(false);
      sandbox.stub(pathUtils, "getEnvFilePath").resolves(ok("."));
      sandbox.stub(settingsUtil, "readSettings").resolves(ok(mockSettings));
      const res = await envUtil.readEnv(".", "dev", false, false);
      assert.isTrue(res.isErr());
      if (res.isErr()) {
        assert.isTrue(res.error instanceof FileNotFoundError);
      }
    });
    it("not silent 2", async () => {
      sandbox.stub(fs, "pathExists").resolves(false);
      sandbox.stub(pathUtils, "getEnvFilePath").resolves(ok(""));
      sandbox.stub(settingsUtil, "readSettings").resolves(ok(mockSettings));
      const res = await envUtil.readEnv(".", "dev", false, false);
      assert.isTrue(res.isErr());
      if (res.isErr()) {
        assert.isTrue(res.error instanceof FileNotFoundError);
      }
    });

    it("loadToProcessEnv false", async () => {
      sandbox.stub(pathUtils, "getEnvFilePath").resolves(ok(".env.dev"));
      const encRes = await cryptoProvider.encrypt(decrypted);
      if (encRes.isErr()) throw encRes.error;
      const encrypted = encRes.value;
      sandbox.stub(fs, "readFile").resolves(("SECRET_ABC=" + encrypted) as any);
      sandbox.stub(fs, "pathExists").resolves(true);
      sandbox.stub(settingsUtil, "readSettings").resolves(ok(mockSettings));
      const res = await envUtil.readEnv(".", "dev", false);
      assert.isTrue(res.isOk());
      assert.equal(process.env.SECRET_ABC, decrypted);
    });

    it("read settings.json fail", async () => {
      sandbox.stub(pathUtils, "getEnvFilePath").resolves(ok(".env.dev"));
      sandbox.stub(fs, "readFile").resolves("SECRET_ABC=AAA" as any);
      sandbox.stub(fs, "pathExists").resolves(true);
      sandbox
        .stub(settingsUtil, "readSettings")
        .resolves(err(new UserError({ source: "test", name: "TestError", message: "message" })));
      const res = await envUtil.readEnv(".", "dev");
      assert.isTrue(res.isErr());
    });
  });

  it("MissingEnvironmentVariablesError", async () => {
    const error1 = new MissingEnvironmentVariablesError("test", "ABC", "./abc.yml");
    assert.isTrue(error1.message.includes("abc.yml"));
    globalVars.ymlFilePath = "./abc.yml";
    const error2 = new MissingEnvironmentVariablesError("test", "ABC");
    assert.isTrue(error2.message.includes("abc.yml"));
    globalVars.ymlFilePath = "";
    const error3 = new MissingEnvironmentVariablesError("test", "ABC");
    assert.isFalse(error3.message.includes("abc.yml"));
  });

  describe("pathUtils.writeEnv", () => {
    beforeEach(() => {
      sandbox.stub(fs, "ensureFile").resolves();
    });
    afterEach(() => {
      sandbox.restore();
    });
    it("happy path", async () => {
      sandbox.stub(pathUtils, "getEnvFilePath").resolves(ok(".env.dev"));
      let value = "";
      sandbox.stub(fs, "writeFile").callsFake(async (file: fs.PathLike | number, data: any) => {
        value = data as string;
        return Promise.resolve();
      });
      sandbox.stub(settingsUtil, "readSettings").resolves(ok(mockSettings));
      const res = await envUtil.writeEnv(".", "dev", { SECRET_ABC: decrypted });
      assert.isTrue(res.isOk());
      assert.isDefined(value);
      value = value!.substring("SECRET_ABC=".length);
      const decRes = await cryptoProvider.decrypt(value);
      if (decRes.isErr()) throw decRes.error;
      assert.isTrue(decRes.isOk());
      assert.equal(decRes.value, decrypted);
    });
    it("no variables", async () => {
      sandbox.stub(pathUtils, "getEnvFilePath").resolves(ok(".env.dev"));
      sandbox.stub(fs, "readFile").resolves("" as any);
      sandbox.stub(fs, "writeFile").resolves();
      sandbox.stub(fs, "pathExists").resolves(true);
      sandbox.stub(settingsUtil, "readSettings").resolves(ok(mockSettings));
      const res = await envUtil.writeEnv(".", "dev", {});
      assert.isTrue(res.isOk());
    });
    it("write to default path", async () => {
      sandbox.stub(pathUtils, "getEnvFilePath").resolves(ok(undefined));
      sandbox.stub(settingsUtil, "readSettings").resolves(ok(mockSettings));
      sandbox.stub(fs, "writeFile").resolves();
      const res = await envUtil.writeEnv(".", "dev", {
        SECRET_ABC: decrypted,
        TEAMS_APP_UPDATE_TIME: "xx-xx-xx",
      });
      assert.isTrue(res.isOk());
    });
    it("write failed", async () => {
      sandbox.stub(pathUtils, "getEnvFilePath").resolves(ok(".env.dev"));
      sandbox
        .stub(settingsUtil, "readSettings")
        .resolves(err(new UserError({ source: "test", name: "TestError", message: "message" })));
      const res = await envUtil.writeEnv(".", "dev", { SECRET_ABC: decrypted });
      assert.isTrue(res.isErr());
    });
  });

  describe("pathUtils.listEnv", () => {
    it("happy path", async () => {
      sandbox.stub(pathUtils, "getEnvFolderPath").resolves(ok("teamsfx"));
      sandbox.stub(fs, "readdir").resolves([".env.dev", ".env.prod"] as any);
      const res = await envUtil.listEnv(".");
      assert.isTrue(res.isOk());
      if (res.isOk()) {
        assert.deepEqual(res.value, ["dev", "prod"]);
      }
    });

    it("remote env only", async () => {
      sandbox.stub(pathUtils, "getEnvFolderPath").resolves(ok("teamsfx"));
      sandbox
        .stub(fs, "readdir")
        .resolves([".env.dev", ".env.prod", ".env.local", ".env.testtool"] as any);
      const res = await envUtil.listEnv(".", true);
      assert.isTrue(res.isOk());
      if (res.isOk()) {
        assert.deepEqual(res.value, ["dev", "prod"]);
      }
    });
  });

  describe("pathUtils.mergeEnv", () => {
    it("case 1", async () => {
      const env: DotenvOutput = {};
      mockedEnvRestore = mockedEnv({
        mykey: "myvalue",
      });
      envUtil.mergeEnv(process.env, env);
      assert.equal(process.env.mykey, "myvalue");
    });

    it("case 2", async () => {
      const env: DotenvOutput = { mykey: "myvalue" };
      mockedEnvRestore = mockedEnv({
        mykey: "",
      });
      envUtil.mergeEnv(process.env, env);
      assert.equal(process.env.mykey, "myvalue");
    });

    it("case 3", async () => {
      const env: DotenvOutput = { mykey: "myvalue2" };
      mockedEnvRestore = mockedEnv({
        mykey: "myvalue",
      });
      envUtil.mergeEnv(process.env, env);
      assert.equal(process.env.mykey, "myvalue");
    });

    it("case 4", async () => {
      const env: DotenvOutput = { mykey: "" };
      mockedEnvRestore = mockedEnv({
        mykey: "myvalue",
      });
      envUtil.mergeEnv(process.env, env);
      assert.equal(process.env.mykey, "myvalue");
    });
  });

  describe("environmentManager", () => {
    it("environmentManager.listAllEnvConfigs", async () => {
      sandbox.stub(pathUtils, "getEnvFolderPath").resolves(ok("teamsfx"));
      sandbox.stub(fs, "readdir").resolves([".env.dev", ".env.prod"] as any);
      const res = await environmentManager.listAllEnvConfigs(".");
      assert.isTrue(res.isOk());
      if (res.isOk()) {
        assert.deepEqual(res.value, ["dev", "prod"]);
      }
    });
    it("environmentManager.listAllEnvConfigs projectPath doesn't exist", async () => {
      sandbox.stub(fs, "pathExists").resolves(false);
      const res = await environmentManager.listAllEnvConfigs(".");
      assert.isFalse(res.isOk());
      assert.instanceOf(res._unsafeUnwrapErr(), FileNotFoundError);
    });
    it("environmentManager.listRemoteEnvConfigs", async () => {
      sandbox.stub(pathUtils, "getEnvFolderPath").resolves(ok("teamsfx"));
      sandbox.stub(fs, "readdir").resolves([".env.dev", ".env.prod", ".env.local"] as any);
      const res = await environmentManager.listRemoteEnvConfigs(".");
      assert.isTrue(res.isOk());
      if (res.isOk()) {
        assert.deepEqual(res.value, ["dev", "prod"]);
      }
    });
    it("environmentManager.listRemoteEnvConfigs projectPath doesn't exist", async () => {
      sandbox.stub(fs, "pathExists").resolves(false);
      const res = await environmentManager.listRemoteEnvConfigs(".");
      assert.isFalse(res.isOk());
      assert.instanceOf(res._unsafeUnwrapErr(), FileNotFoundError);
    });
    it("environmentManager.listRemoteEnvConfigs no remote env, only local", async () => {
      sandbox.stub(pathUtils, "getEnvFolderPath").resolves(ok("teamsfx"));
      sandbox.stub(fs, "readdir").resolves([".env.local"] as any);
      const res = await environmentManager.listRemoteEnvConfigs(".", true);
      assert.isFalse(res.isOk());
      assert.instanceOf(res._unsafeUnwrapErr(), NoEnvFilesError);
    });
    it("environmentManager.listRemoteEnvConfigs return error", async () => {
      sandbox.stub(fs, "pathExists").resolves(false);
      sandbox.stub(fs, "readdir").resolves([] as any);
      sandbox.stub(pathUtils, "getYmlFilePath").returns("./xxx");
      const res = await environmentManager.listRemoteEnvConfigs(".", true);
      assert.isTrue(res.isErr());
    });
    it("environmentManager.getExistingNonRemoteEnvs with testtool env", async () => {
      sandbox.stub(pathUtils, "getEnvFolderPath").resolves(ok("teamsfx"));
      sandbox
        .stub(fs, "readdir")
        .resolves([".env.dev", ".env.prod", ".env.local", ".env.testtool"] as any);
      const res = await environmentManager.getExistingNonRemoteEnvs(".");
      assert.deepEqual(res, ["testtool", "local"]);
    });
    it("environmentManager.getExistingNonRemoteEnvs without testtool env", async () => {
      sandbox.stub(pathUtils, "getEnvFolderPath").resolves(ok("teamsfx"));
      sandbox.stub(fs, "readdir").resolves([".env.dev", ".env.prod", ".env.local"] as any);
      const res = await environmentManager.getExistingNonRemoteEnvs(".");
      assert.deepEqual(res, ["local"]);
    });
    it("environmentManager.getExistingNonRemoteEnvs without projectPath", async () => {
      sandbox.stub(fs, "pathExists").resolves(false);
      const res = await environmentManager.getExistingNonRemoteEnvs(".");
      assert.deepEqual(res, ["local"]);
    });
  });

  describe("EnvLoaderMW", () => {
    it("EnvLoaderMW success", async () => {
      sandbox.stub(pathUtils, "getEnvFolderPath").resolves(ok("teamsfx"));
      const encRes = await cryptoProvider.encrypt(decrypted);
      if (encRes.isErr()) throw encRes.error;
      const encrypted = encRes.value;
      sandbox.stub(fs, "pathExists").resolves(true);
      sandbox.stub(fs, "readFile").resolves(("SECRET_ABC=" + encrypted) as any);
      sandbox.stub(settingsUtil, "readSettings").resolves(ok(mockSettings));
      if (process.env.SECRET_ABC || process.env.SECRET_ABC === undefined) {
        delete process.env.SECRET_ABC;
      }
      sandbox
        .stub(dotenvUtil, "deserialize")
        .onFirstCall()
        .returns({
          lines: [],
          obj: {},
        })
        .onSecondCall()
        .returns({
          lines: [],
          obj: { SECRET_ABC: encrypted },
        })
        .onThirdCall()
        .returns({
          lines: [],
          obj: {},
        });
      process.env.ENV_VAR = "1";
      class MyClass {
        async myMethod(inputs: Inputs): Promise<Result<any, FxError>> {
          assert.equal(process.env.SECRET_ABC, decrypted);
          process.env.ENV_VAR = "2";
          return ok(undefined);
        }
      }
      hooks(MyClass, {
        myMethod: [EnvLoaderMW(true)],
      });
      const my = new MyClass();
      const inputs = {
        platform: Platform.VSCode,
        projectPath: ".",
        env: "dev",
      };
      const res = await my.myMethod(inputs);
      assert.isTrue(res.isOk());
      assert.isUndefined(process.env.SECRET_ABC);
      assert.equal(process.env.ENV_VAR, "1", "process.env.ENV_VAR should be restored to 1");
    });

    it("EnvLoaderMW skip load", async () => {
      sandbox.stub(fs, "pathExists").resolves(true);
      class MyClass {
        async myMethod(inputs: Inputs): Promise<Result<any, FxError>> {
          return ok(undefined);
        }
      }
      hooks(MyClass, {
        myMethod: [EnvLoaderMW(true, true)],
      });
      const my = new MyClass();
      const inputs = {
        platform: Platform.VSCode,
        projectPath: ".",
      };
      const res = await my.myMethod(inputs);
      assert.isTrue(res.isOk());
    });

    it("EnvLoaderMW ignore-env-file", async () => {
      sandbox.stub(fs, "pathExists").resolves(true);
      class MyClass {
        async myMethod(inputs: Inputs): Promise<Result<any, FxError>> {
          return ok(undefined);
        }
      }
      hooks(MyClass, {
        myMethod: [EnvLoaderMW(true)],
      });
      const my = new MyClass();
      const inputs = {
        platform: Platform.VSCode,
        projectPath: ".",
        "ignore-env-file": true,
      };
      const res = await my.myMethod(inputs);
      assert.isTrue(res.isOk());
    });

    it("EnvLoaderMW success for F5 (missing .env file)", async () => {
      sandbox.stub(pathUtils, "getEnvFilePath").resolves(ok("."));
      sandbox.stub(fs, "pathExistsSync").returns(false);
      sandbox.stub(fs, "writeFile").resolves();
      sandbox.stub(envUtil, "readEnv").resolves(ok({}));
      class MyClass {
        async myMethod(inputs: Inputs): Promise<Result<any, FxError>> {
          return ok(undefined);
        }
      }
      hooks(MyClass, {
        myMethod: [EnvLoaderMW(false)],
      });
      const my = new MyClass();
      const inputs = {
        platform: Platform.VSCode,
        projectPath: ".",
        env: "dev",
        isLocalDebug: true,
      };
      const res = await my.myMethod(inputs);
      assert.isTrue(res.isOk());
    });
    it("EnvLoaderMW failed for F5 (missing .env file and getEnvFilePath Error)", async () => {
      sandbox.stub(pathUtils, "getEnvFilePath").resolves(err(new UserError({})));
      class MyClass {
        async myMethod(inputs: Inputs): Promise<Result<any, FxError>> {
          return ok(undefined);
        }
      }
      hooks(MyClass, {
        myMethod: [EnvLoaderMW(false)],
      });
      const my = new MyClass();
      const inputs = {
        platform: Platform.VSCode,
        projectPath: ".",
        env: "dev",
        isLocalDebug: true,
      };
      const res = await my.myMethod(inputs);
      assert.isTrue(res.isErr());
    });
    it("EnvLoaderMW success: no env available", async () => {
      sandbox.stub(pathUtils, "getEnvFolderPath").resolves(ok("teamsfx"));
      sandbox.stub(envUtil, "listEnv").resolves(ok([]));
      sandbox.stub(envUtil, "readEnv").resolves(ok({}));
      class MyClass {
        async myMethod(inputs: Inputs): Promise<Result<any, FxError>> {
          return ok(undefined);
        }
      }
      hooks(MyClass, {
        myMethod: [EnvLoaderMW(true)],
      });
      const my = new MyClass();
      const inputs = {
        platform: Platform.VSCode,
        projectPath: ".",
      };
      const res = await my.myMethod(inputs);
      assert.isTrue(res.isOk());
    });
    it("EnvLoaderMW ignoreEnvInfo", async () => {
      sandbox.stub(pathUtils, "getEnvFolderPath").resolves(ok("teamsfx"));
      sandbox.stub(envUtil, "readEnv").resolves(ok({}));
      class MyClass {
        async myMethod(inputs: Inputs): Promise<Result<any, FxError>> {
          return ok(undefined);
        }
      }
      hooks(MyClass, {
        myMethod: [EnvLoaderMW(true)],
      });
      const my = new MyClass();
      const inputs = {
        platform: Platform.VSCode,
        projectPath: ".",
        ignoreEnvInfo: true,
      };
      const res = await my.myMethod(inputs);
      assert.isTrue(res.isOk());
      const core = new FxCore(tools);
    });
    it("EnvLoaderMW fail without projectPath", async () => {
      class MyClass {
        async myMethod(inputs: Inputs): Promise<Result<any, FxError>> {
          return ok(undefined);
        }
      }
      hooks(MyClass, {
        myMethod: [EnvLoaderMW(true)],
      });
      const my = new MyClass();
      const inputs = {
        platform: Platform.VSCode,
        env: "dev",
      };
      const res = await my.myMethod(inputs);
      assert.isTrue(res.isErr());
      if (res.isErr()) {
        assert.equal(res.error.name, "NoProjectOpenedError");
      }
    });
    it("EnvLoaderMW fail with listEnv Error", async () => {
      sandbox.stub(pathUtils, "getEnvFolderPath").resolves(ok("teamsfx"));
      // sandbox
      //   .stub(envUtil, "listEnv")
      //   .resolves(err(new UserError({ source: "test", name: "TestError", message: "message" })));
      class MyClass {
        async myMethod(inputs: Inputs): Promise<Result<any, FxError>> {
          return ok(undefined);
        }
      }
      sandbox.stub(TOOLS.ui, "selectOption").resolves(err(new UserError({})));
      hooks(MyClass, {
        myMethod: [EnvLoaderMW(true)],
      });
      const my = new MyClass();
      const inputs = {
        platform: Platform.VSCode,
        projectPath: ".",
      };
      const res = await my.myMethod(inputs);
      assert.isTrue(res.isErr());
    });
    it("EnvLoaderMW fail with envUtil Error", async () => {
      const encRes = await cryptoProvider.encrypt(decrypted);
      if (encRes.isErr()) throw encRes.error;
      const encrypted = encRes.value;
      sandbox.stub(fs, "pathExists").resolves(true);
      sandbox.stub(fs, "readFile").resolves(("SECRET_ABC=" + encrypted) as any);
      sandbox.stub(settingsUtil, "readSettings").resolves(ok(mockSettings));
      sandbox
        .stub(envUtil, "readEnv")
        .resolves(err(new UserError({ source: "test", name: "TestError", message: "message" })));
      class MyClass {
        async myMethod(inputs: Inputs): Promise<Result<any, FxError>> {
          return ok(undefined);
        }
      }
      hooks(MyClass, {
        myMethod: [EnvLoaderMW(true)],
      });
      const my = new MyClass();
      const inputs = {
        platform: Platform.VSCode,
        projectPath: ".",
        env: "dev",
      };
      const res = await my.myMethod(inputs);
      assert.isTrue(res.isErr());
    });
    it("EnvLoaderMW cancel", async () => {
      sandbox.stub(envUtil, "listEnv").resolves(ok(["dev", "prod"]));
      sandbox.stub(tools.ui, "selectOption").resolves(err(new UserCancelError()));
      class MyClass {
        async myMethod(inputs: Inputs): Promise<Result<any, FxError>> {
          return ok(undefined);
        }
      }
      hooks(MyClass, {
        myMethod: [EnvLoaderMW(true)],
      });
      const my = new MyClass();
      const inputs = {
        platform: Platform.VSCode,
        projectPath: ".",
      };
      const res = await my.myMethod(inputs);
      assert.isTrue(res.isErr());
    });
  });

  describe("EnvWriterMW", () => {
    beforeEach(() => {
      sandbox.stub(fs, "ensureFile").resolves();
    });
    afterEach(() => {
      sandbox.restore();
    });
    it("EnvWriterMW success", async () => {
      sandbox.stub(pathUtils, "getEnvFolderPath").resolves(ok("teamsfx"));
      let value = "";
      sandbox.stub(fs, "writeFile").callsFake(async (file: fs.PathLike | number, data: any) => {
        value = data as string;
        return Promise.resolve();
      });
      sandbox.stub(settingsUtil, "readSettings").resolves(ok(mockSettings));
      const envs = { SECRET_ABC: decrypted };
      class MyClass {
        async myMethod(inputs: Inputs, ctx?: CoreHookContext): Promise<Result<any, FxError>> {
          ctx!.envVars = envs;
          return ok(undefined);
        }
      }
      hooks(MyClass, {
        myMethod: [ContextInjectorMW, EnvWriterMW],
      });
      const my = new MyClass();
      const inputs = {
        platform: Platform.VSCode,
        projectPath: ".",
        env: "dev",
      };
      const res = await my.myMethod(inputs);
      assert.isTrue(res.isOk());
      assert.isDefined(value);
      value = value!.substring("SECRET_ABC=".length);
      const decRes = await cryptoProvider.decrypt(value);
      if (decRes.isErr()) throw decRes.error;
      assert.isTrue(decRes.isOk());
      assert.equal(decRes.value, decrypted);
    });

    it("EnvWriterMW fail with envUtil Error", async () => {
      sandbox
        .stub(envUtil, "writeEnv")
        .resolves(err(new UserError({ source: "test", name: "TestError", message: "message" })));
      sandbox.stub(settingsUtil, "readSettings").resolves(ok(mockSettings));
      const envs = { SECRET_ABC: decrypted };
      class MyClass {
        async myMethod(inputs: Inputs, ctx?: CoreHookContext): Promise<Result<any, FxError>> {
          ctx!.envVars = envs;
          return ok(undefined);
        }
      }
      hooks(MyClass, {
        myMethod: [ContextInjectorMW, EnvWriterMW],
      });
      const my = new MyClass();
      const inputs = {
        platform: Platform.VSCode,
        projectPath: ".",
        env: "dev",
      };
      const res = await my.myMethod(inputs);
      assert.isTrue(res.isErr());
    });
  });

  describe("dotenvUtil", () => {
    it("dotenvUtil deserialize & serialize", async () => {
      const original =
        '#COMMENT\n\n\nKEY1=VALUE1#COMMENT2\nKEY2=\'VALUE2\'\nKEY3="VALUE3#"\nindexPath="/index.html#"#COMMENT3';
      const expected =
        '#COMMENT\n\n\nKEY1=VALUE1 #COMMENT2\nKEY2=\'VALUE2\'\nKEY3="VALUE3#"\nindexPath="/index.html#" #COMMENT3\nKEY4="VALUE4"\nKEY5="VALUE5#"';
      const parsed = dotenvUtil.deserialize(original);
      console.log(parsed);
      assert.deepEqual(parsed, {
        lines: [
          "#COMMENT",
          "",
          "",
          { key: "KEY1", value: "VALUE1", comment: "#COMMENT2" },
          { key: "KEY2", value: "VALUE2", quote: "'" },
          { key: "KEY3", value: "VALUE3#", quote: '"' },
          { key: "indexPath", value: "/index.html#", quote: '"', comment: "#COMMENT3" },
        ],
        obj: { KEY1: "VALUE1", KEY2: "VALUE2", KEY3: "VALUE3#", indexPath: "/index.html#" },
      });
      parsed.lines?.push({ key: "KEY4", value: "VALUE4", quote: '"' });
      parsed.obj["KEY5"] = "VALUE5#";
      const serialized = dotenvUtil.serialize(parsed);
      assert.equal(serialized, expected);
    });
    it("dotenvUtil deserialize & serialize empty", async () => {
      const original = "";
      const parsed = dotenvUtil.deserialize(original);
      assert.deepEqual(parsed, {
        lines: [""],
        obj: {},
      });
      const serialized = dotenvUtil.serialize(parsed);
      assert.equal(serialized, original);
    });
    it("dotenvUtil serialize without lines", async () => {
      const parsed = {
        obj: { KEY: "VALUE", KEY2: "VALUE2" },
      };
      const str = dotenvUtil.serialize(parsed);
      assert.equal(str, "KEY=VALUE\nKEY2=VALUE2");
    });
  });

  describe("settingsUtil", () => {
    afterEach(() => {
      sandbox.restore();
    });
    it("settingsUtil read not exist", async () => {
      sandbox.stub(fs, "pathExists").resolves(false);
      sandbox.stub(pathUtils, "getYmlFilePath").returns(".");
      const res = await settingsUtil.readSettings("abc");
      assert.isTrue(res.isErr());
    });

    it("settingsUtil read and ensure trackingId", async () => {
      sandbox.stub(fs, "pathExists").resolves(true);
      sandbox.stub(pathUtils, "getYmlFilePath").returns(".");
      sandbox.stub(fs, "readFile").resolves("version: 1.0.0" as any);
      sandbox.stub(fs, "writeFile").resolves();
      const res = await settingsUtil.readSettings("abc");
      assert.isTrue(res.isOk());
      if (res.isOk()) {
        assert.isDefined(res.value.trackingId);
      }
    });

    it("settingsUtil write success", async () => {
      sandbox.stub(pathUtils, "getYmlFilePath").returns(".");
      sandbox.stub(fs, "writeFile").resolves();
      sandbox.stub(fs, "pathExists").resolves(true);
      sandbox.stub(fs, "readFile").resolves("version: 1.0.0" as any);
      const res = await settingsUtil.writeSettings(".", { trackingId: "123", version: "2" });
      assert.isTrue(res.isOk());
    });
    it("settingsUtil write failed", async () => {
      sandbox.stub(pathUtils, "getYmlFilePath").returns(".");
      sandbox.stub(fs, "pathExists").resolves(false);
      const res = await settingsUtil.writeSettings(".", { trackingId: "123", version: "2" });
      assert.isTrue(res.isErr());
      assert.isTrue(res._unsafeUnwrapErr() instanceof FileNotFoundError);
    });
  });

  describe("extractEnvNameFromFileName", () => {
    it("happy path", async () => {
      const res = await envUtil.extractEnvNameFromFileName(".env.dev");
      assert.isTrue(res === "dev");
    });
    it("return undefined", async () => {
      const res = await envUtil.extractEnvNameFromFileName(".env.dev.user");
      assert.isUndefined(res);
    });
    it("return undefined", async () => {
      const res = await envUtil.extractEnvNameFromFileName(".env1.dev");
      assert.isTrue(res === undefined);
    });
  });

  describe("loadEnvFile", () => {
    it("happy path", async () => {
      sandbox.stub(dotenvUtil, "deserialize").returns({ obj: { KEY: "VALUE" } });
      sandbox.stub(fs, "readFile").resolves("" as any);
      await envUtil.loadEnvFile(".env.dev");
      assert.equal(process.env.KEY, "VALUE");
    });
  });

  describe("resetEnvFile", () => {
    it("happy path", async () => {
      const obj: any = { obj: { IKEY: "IKEY", KEY: "KEY" } };
      sandbox.stub(dotenvUtil, "deserialize").returns(obj);
      sandbox.stub(fs, "readFile").resolves("" as any);
      sandbox.stub(fs, "writeFile").resolves();
      await envUtil.resetEnvFile(" ", ["IKEY"]);
      assert.equal(obj.obj.KEY, "");
    });
  });

  describe("resetEnv", () => {
    it("happy path", async () => {
      sandbox.stub(pathUtils, "getEnvFilePath").resolves(ok(".env.dev"));
      sandbox.stub(envUtil, "resetEnvFile").resolves();
      sandbox.stub(fs, "pathExists").resolves(true);
      await envUtil.resetEnv(" ", "dev", ["IKEY"]);
    });
    it("getEnvFilePath error", async () => {
      sandbox.stub(pathUtils, "getEnvFilePath").resolves(err(new UserCancelError()));
      await envUtil.resetEnv(" ", "dev", ["IKEY"]);
    });
    it("getEnvFilePath return undefined", async () => {
      sandbox.stub(pathUtils, "getEnvFilePath").resolves(ok(undefined));
      await envUtil.resetEnv(" ", "dev", ["IKEY"]);
    });
  });
});

describe("parseSetOutputCommand", () => {
  const tools = new MockTools();
  setTools(tools);
  const sandbox = sinon.createSandbox();
  let mockedEnvRestore: RestoreFn | undefined;
  afterEach(() => {
    sandbox.restore();
    if (mockedEnvRestore) {
      mockedEnvRestore();
    }
  });
  it("parse one key value pair", async () => {
    const res = parseSetOutputCommand('echo "::set-teamsfx-env TAB_DOMAIN=localhost:53000"');
    assert.deepEqual(res, { TAB_DOMAIN: "localhost:53000" });
  });
  it("parse two key value pairs", async () => {
    const res = parseSetOutputCommand(
      'echo "::set-teamsfx-env TAB_DOMAIN=localhost:53000"; echo "::set-teamsfx-env TAB_ENDPOINT=https://localhost:53000";'
    );
    assert.deepEqual(res, {
      TAB_DOMAIN: "localhost:53000",
      TAB_ENDPOINT: "https://localhost:53000",
    });
  });
});
