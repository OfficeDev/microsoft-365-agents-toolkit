// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { ConfigFolderName, err, ok } from "@microsoft/teamsfx-api";
import { assert } from "chai";
import "mocha";
import * as nodeFetch from "node-fetch";
import { Response } from "node-fetch";
import os from "os";
import * as sinon from "sinon";
import stream, { Readable } from "stream";
import { getLocalizedString } from "../../../../src/common/localizeUtils";
import { NodeChecker } from "../../../../src/component/deps-checker/internal/nodeChecker";
import { httpClient } from "../../../../src/component/driver/devTool/httpClient";
import {
  NodeDownloadMirror,
  nodejsInstaller,
} from "../../../../src/component/driver/devTool/nodeInstaller";
import { InstallNodeJSError } from "../../../../src/error/depCheck";
import { MockedLogProvider, MockedUserInteraction } from "../../../plugins/solution/util";
import fs from "fs-extra";
import path from "path";
import { UserCancelError } from "../../../../src/error";

describe("NodeJS Installer", () => {
  const sandbox = sinon.createSandbox();

  describe("HttpClient", () => {
    afterEach(() => {
      sandbox.restore();
    });

    describe("get", () => {
      it("fetch return 500", async () => {
        sandbox.stub(nodeFetch, "default").resolves({ ok: false, status: 500 } as any);
        try {
          await httpClient.get("https://test.com");
        } catch (e: any) {
          assert.equal(e.message, "Request failed with status 500");
        }
      });

      it("happy", async () => {
        const buffer = Buffer.from("chunk1");
        const fakeResponse = new Response(Readable.from(buffer), {
          status: 200,
          headers: { "content-type": "application/json" },
        });
        sandbox.stub(nodeFetch, "default").resolves(fakeResponse);
        const result = await httpClient.get("https://test.com", { progress: () => {} });
        assert.equal(result.toString(), "chunk1");
      });
    });

    it("getText", async () => {
      sandbox.stub(httpClient, "get").resolves(Buffer.from("chunk1chunk2"));
      const result = await httpClient.getText("https://test.com");
      assert.equal(result, "chunk1chunk2");
    });

    describe("headTime", () => {
      it("fetch return 500", async () => {
        sandbox.stub(nodeFetch, "default").resolves({ ok: false, status: 500 } as any);
        try {
          await httpClient.headTime("https://test.com");
        } catch (e: any) {
          assert.equal(e.message, "Request failed with status 500");
        }
      });

      it("happy", async () => {
        const fakeResponse = new Response(undefined, {
          status: 200,
          headers: { "content-type": "application/json" },
        });
        sandbox.stub(nodeFetch, "default").resolves(fakeResponse);
        const result = await httpClient.headTime("https://test.com");
        assert.isDefined(result);
      });
    });
  });

  describe("NodejsInstaller", () => {
    afterEach(() => {
      sandbox.restore();
    });

    describe("getNameAndExt", () => {
      it("darwin-arm64", async () => {
        sandbox.stub(os, "platform").returns("darwin");
        sandbox.stub(os, "arch").returns("arm64");
        const { name, ext } = nodejsInstaller.getNameAndExt();
        assert.equal(name, "darwin-arm64");
        assert.equal(ext, ".tar.xz");
      });
      it("linux-x64", async () => {
        sandbox.stub(os, "platform").returns("linux");
        sandbox.stub(os, "arch").returns("x64");
        const { name, ext } = nodejsInstaller.getNameAndExt();
        assert.equal(name, "linux-x64");
        assert.equal(ext, ".tar.xz");
      });
      it("win-x64", async () => {
        sandbox.stub(os, "platform").returns("win32");
        sandbox.stub(os, "arch").returns("x64");
        const { name, ext } = nodejsInstaller.getNameAndExt();
        assert.equal(name, "win-x64");
        assert.equal(ext, ".zip");
      });
      it("aix-x64", async () => {
        sandbox.stub(os, "platform").returns("aix");
        sandbox.stub(os, "arch").returns("x64");
        const { name, ext } = nodejsInstaller.getNameAndExt();
        assert.equal(name, "aix-x64");
        assert.equal(ext, ".tar.gz");
      });
    });
  });

  describe("getLatestLTSVersion", () => {
    it("happy", async () => {
      const mirror: NodeDownloadMirror = {
        url: "https://nodejs.org/dist",
        name: "test mirror",
        indexJsonUrl: "https://nodejs.org/dist/index.json",
        indexJson: [
          {
            version: "v23.0.0",
            lts: false,
          },
          {
            version: "v22.0.0",
            lts: "Argon",
          },
        ],
      };
      const version = nodejsInstaller.getLatestLTSVersion(mirror);
      assert.equal(version, "v22.0.0");
    });

    it("LTS not found", async () => {
      const mirror: NodeDownloadMirror = {
        url: "https://nodejs.org/dist",
        name: "test mirror",
        indexJsonUrl: "https://nodejs.org/dist/index.json",
        indexJson: [
          {
            version: "v23.0.0",
            lts: false,
          },
          {
            version: "v22.0.0",
            lts: false,
          },
        ],
      };
      const version = nodejsInstaller.getLatestLTSVersion(mirror);
      assert.isUndefined(version);
    });
  });

  describe("fetchJSON", () => {
    afterEach(() => {
      sandbox.restore();
    });
    it("happy", async () => {
      sandbox.stub(httpClient, "getText").resolves(JSON.stringify({ version: "v22.0.0" }));
      const jsonRes = await nodejsInstaller.fetchJSON("test url");
      assert.isTrue(jsonRes.isOk());
      if (jsonRes.isOk()) {
        assert.deepEqual(jsonRes.value, { version: "v22.0.0" });
      }
    });

    it("error", async () => {
      sandbox.stub(httpClient, "getText").rejects(new Error("test error"));
      const jsonRes = await nodejsInstaller.fetchJSON("test url");
      assert.isTrue(jsonRes.isErr());
      if (jsonRes.isErr()) {
        assert.isTrue(jsonRes.error instanceof InstallNodeJSError);
      }
    });
  });

  describe("fetchString", () => {
    afterEach(() => {
      sandbox.restore();
    });
    it("happy", async () => {
      sandbox.stub(httpClient, "getText").resolves("abcd");
      const textRes = await nodejsInstaller.fetchString("test url");
      assert.isTrue(textRes.isOk());
      if (textRes.isOk()) {
        assert.deepEqual(textRes.value, "abcd");
      }
    });

    it("error", async () => {
      sandbox.stub(httpClient, "getText").rejects(new Error("test error"));
      const textRes = await nodejsInstaller.fetchString("test url");
      assert.isTrue(textRes.isErr());
      if (textRes.isErr()) {
        assert.isTrue(textRes.error instanceof InstallNodeJSError);
      }
    });
  });

  describe("fetchBinary", () => {
    afterEach(() => {
      sandbox.restore();
    });
    it("happy", async () => {
      sandbox.stub(httpClient, "get").resolves(Buffer.from("abcd"));
      const binRes = await nodejsInstaller.fetchBinary("test url");
      assert.isTrue(binRes.isOk());
      if (binRes.isOk()) {
        assert.deepEqual(binRes.value.toString(), "abcd");
      }
    });

    it("error", async () => {
      sandbox.stub(httpClient, "get").rejects(new Error("test error"));
      const binRes = await nodejsInstaller.fetchString("test url");
      assert.isTrue(binRes.isErr());
      if (binRes.isErr()) {
        assert.isTrue(binRes.error instanceof InstallNodeJSError);
      }
    });
  });

  describe("resolveUrl", () => {
    it("absolute url", async () => {
      const url = nodejsInstaller.resolveUrl(
        "https://registry.npmmirror.com/-/binary/node/",
        "https://registry.npmmirror.com/-/binary/node/index.json"
      );
      assert.equal(url, "https://registry.npmmirror.com/-/binary/node/index.json");
    });
    it("relative to base url, target is folder", async () => {
      const url = nodejsInstaller.resolveUrl("https://nodejs.org/dist/", "v0.11.7/");
      assert.equal(url, "https://nodejs.org/dist/v0.11.7/");
    });
    it("relative to base url, target is a file", async () => {
      const url = nodejsInstaller.resolveUrl("https://example.com/path/to/page.html", "image.jpg");
      assert.equal(url, "https://example.com/path/to/image.jpg");
    });
    it("relative to domain", async () => {
      const url = nodejsInstaller.resolveUrl(
        "https://nodejs.org/dist/v22.14.0/",
        "/dist/latest-v22.x/node-v22.14.0-linux-armv7l.tar.gz"
      );
      assert.equal(url, "https://nodejs.org/dist/latest-v22.x/node-v22.14.0-linux-armv7l.tar.gz");
    });
  });

  describe("testMirrorSpeed", () => {
    afterEach(() => {
      sandbox.restore();
    });
    it("error", async () => {
      sandbox.stub(httpClient, "headTime").rejects(new Error("test error"));
      const mirror: NodeDownloadMirror = {
        url: "https://nodejs.org/dist",
        name: "test mirror",
        indexJsonUrl: "https://nodejs.org/dist/index.json",
      };
      const res = await nodejsInstaller.testMirrorSpeed(mirror, "win-x64", ".zip", 1000);
      assert.isUndefined(res.indexJson);
    });
    it("no lts version", async () => {
      sandbox.stub(httpClient, "headTime").resolves(1000);
      sandbox.stub(httpClient, "getText").resolves("[]");
      const mirror: NodeDownloadMirror = {
        url: "https://nodejs.org/dist",
        name: "test mirror",
        indexJsonUrl: "https://nodejs.org/dist/index.json",
      };
      const res = await nodejsInstaller.testMirrorSpeed(mirror, "win-x64", ".zip", 1000);
      assert.deepEqual(res.indexJson, []);
      assert.isUndefined(res.version);
    });
    it("get download url fail", async () => {
      sandbox.stub(httpClient, "headTime").resolves(1000);
      sandbox.stub(httpClient, "getText").resolves("[]");
      sandbox.stub(nodejsInstaller, "getLatestLTSVersion").returns("v22.14.0");
      sandbox
        .stub(nodejsInstaller, "getDownloadUrl")
        .resolves(err(new InstallNodeJSError("test error")));
      const mirror: NodeDownloadMirror = {
        url: "https://nodejs.org/dist",
        name: "test mirror",
        indexJsonUrl: "https://nodejs.org/dist/index.json",
      };
      const res = await nodejsInstaller.testMirrorSpeed(mirror, "win-x64", ".zip", 1000);
      assert.deepEqual(res.indexJson, []);
      assert.isDefined(res.version);
      assert.isUndefined(res.packageUrl);
    });

    it("success", async () => {
      sandbox.stub(httpClient, "headTime").resolves(1000);
      sandbox.stub(httpClient, "getText").resolves("[]");
      sandbox.stub(nodejsInstaller, "getLatestLTSVersion").returns("v22.14.0");
      sandbox
        .stub(nodejsInstaller, "getDownloadUrl")
        .resolves(ok("https://node-v22.14.0-win-x64.zip"));
      const mirror: NodeDownloadMirror = {
        url: "https://nodejs.org/dist",
        name: "test mirror",
        indexJsonUrl: "https://nodejs.org/dist/index.json",
      };
      const res = await nodejsInstaller.testMirrorSpeed(mirror, "win-x64", ".zip", 1000);
      assert.deepEqual(res.indexJson, []);
      assert.equal(res.version, "v22.14.0");
      assert.equal(res.packageUrl, "https://node-v22.14.0-win-x64.zip");
    });
  });

  describe("getBestMirror", () => {
    afterEach(() => {
      sandbox.restore();
    });
    it("FirstPriorityMirror success", async () => {
      const mirror: NodeDownloadMirror = {
        url: "https://nodejs.org/dist",
        name: "test mirror",
        indexJsonUrl: "https://nodejs.org/dist/index.json",
        packageUrl: "https://nodejs.org/dist/v22.14.0/node-v22.14.0-win-x64.zip",
      };
      sandbox.stub(nodejsInstaller, "testMirrorSpeed").resolves(mirror);
      const resultMirror = await nodejsInstaller.getBestMirror("win-x64", ".zip");
      assert.equal(resultMirror, mirror);
    });

    it("FirstPriorityMirror fail, BackupMirrors success", async () => {
      const failMirror: NodeDownloadMirror = {
        url: "https://nodejs.org/dist",
        name: "test mirror",
        indexJsonUrl: "https://nodejs.org/dist/index.json",
      };
      const successMirror: NodeDownloadMirror = {
        url: "https://nodejs.org/dist",
        name: "test mirror",
        indexJsonUrl: "https://nodejs.org/dist/index.json",
        packageUrl: "https://nodejs.org/dist/v22.14.0/node-v22.14.0-win-x64.zip",
      };
      sandbox
        .stub(nodejsInstaller, "testMirrorSpeed")
        .onFirstCall()
        .resolves(failMirror)
        .onSecondCall()
        .resolves(successMirror);
      const resultMirror = await nodejsInstaller.getBestMirror("win-x64", ".zip");
      assert.equal(resultMirror, successMirror);
    });
  });

  describe("parseHtmlToGetUrl", () => {
    afterEach(() => {
      sandbox.restore();
    });
    it("found", async () => {
      const packageUrl = nodejsInstaller.parseHtmlToGetUrl(
        "https://nodejs.org/dist/v22.14.0/",
        '<a href="/dist/v22.14.0/node-v22.14.0-win-x64.zip">node-v22.14.0-win-x64.zip</a>',
        "v22.14.0-win-x64.zip"
      );
      assert.equal(packageUrl, "https://nodejs.org/dist/v22.14.0/node-v22.14.0-win-x64.zip");
    });
    it("not found", async () => {
      const packageUrl = nodejsInstaller.parseHtmlToGetUrl(
        "https://nodejs.org/dist/v22.14.0/",
        '<a href="/dist/v22.14.0/node-v22.14.0-win-x64.zip">node-v22.14.0-win-x64.zip</a>',
        "v22.14.0-linux-x64.zip"
      );
      assert.isUndefined(packageUrl);
    });
  });

  describe("extractPackage", () => {
    afterEach(() => {
      sandbox.restore();
    });
    it("extractZip", async () => {
      sandbox.stub(nodejsInstaller, "getAdmZip").returns({
        extractAllTo: () => {},
      } as any);
      nodejsInstaller.extractZip(Buffer.from(""), "/path/to/dest");
    });
    it("extractTar", async () => {
      sandbox.stub(stream.PassThrough.prototype, "end").returns();
      sandbox.stub(stream.PassThrough.prototype, "pipe").returns({} as any);
      nodejsInstaller.extractTar(Buffer.from(""), "test.tar.gz", "/path/to/dest");
      nodejsInstaller.extractTar(Buffer.from(""), "test.tar.xz", "/path/to/dest");
    });
    it("extractPackage", async () => {
      sandbox.stub(nodejsInstaller, "extractZip").returns();
      sandbox.stub(nodejsInstaller, "extractTar").returns();
      nodejsInstaller.extractPackage(Buffer.from(""), "test.tar.gz", "/path/to/dest");
      nodejsInstaller.extractPackage(Buffer.from(""), "test.zip", "/path/to/dest");
    });
  });

  describe("getDownloadUrl", () => {
    const NpmMirror: NodeDownloadMirror = {
      name: "NPM",
      url: "https://registry.npmmirror.com/-/binary/node/",
      indexJsonUrl: "https://cdn.npmmirror.com/binaries/node/index.json",
    };
    const OfficialMirror: NodeDownloadMirror = {
      name: "Official",
      url: "https://nodejs.org/dist/",
      indexJsonUrl: "https://nodejs.org/dist/index.json",
    };
    afterEach(() => {
      sandbox.restore();
    });
    it("NPM mirror - first fetchJSON fail", async () => {
      sandbox
        .stub(nodejsInstaller, "fetchJSON")
        .resolves(err(new InstallNodeJSError("test error")));
      const res = await nodejsInstaller.getDownloadUrl(NpmMirror, "v22.14.0", "win-x64", ".zip");
      assert.isTrue(res.isErr());
    });
    it("NPM mirror - second fetchJSON fail", async () => {
      sandbox
        .stub(nodejsInstaller, "fetchJSON")
        .onFirstCall()
        .resolves(
          ok([
            {
              name: "v22.14.0",
              url: "https://cdn.npmmirror.com/binaries/node/v22.14.0/",
            },
          ])
        )
        .onSecondCall()
        .resolves(err(new InstallNodeJSError("test error")));
      const res = await nodejsInstaller.getDownloadUrl(NpmMirror, "v22.14.0", "win-x64", ".zip");
      assert.isTrue(res.isErr());
    });
    it("NPM mirror - version not found", async () => {
      sandbox
        .stub(nodejsInstaller, "fetchJSON")
        .onFirstCall()
        .resolves(
          ok([
            {
              name: "v23.14.0",
              url: "https://cdn.npmmirror.com/binaries/node/v22.14.0/",
            },
          ])
        );
      const res = await nodejsInstaller.getDownloadUrl(NpmMirror, "v22.14.0", "win-x64", ".zip");
      assert.isTrue(res.isErr());
    });
    it("NPM mirror - package not found", async () => {
      sandbox
        .stub(nodejsInstaller, "fetchJSON")
        .onFirstCall()
        .resolves(
          ok([
            {
              name: "v22.14.0",
              url: "https://cdn.npmmirror.com/binaries/node/v22.14.0/",
            },
          ])
        )
        .onSecondCall()
        .resolves(ok([]));
      const res = await nodejsInstaller.getDownloadUrl(NpmMirror, "v22.14.0", "win-x64", ".zip");
      assert.isTrue(res.isErr());
    });
    it("NPM mirror - happy", async () => {
      sandbox
        .stub(nodejsInstaller, "fetchJSON")
        .onFirstCall()
        .resolves(
          ok([
            {
              name: "v22.14.0",
              url: "https://cdn.npmmirror.com/binaries/node/v22.14.0/",
            },
          ])
        )
        .onSecondCall()
        .resolves(
          ok([
            {
              name: "v22.14.0-win-x64.zip",
              url: "https://cdn.npmmirror.com/binaries/node/v22.14.0/v22.14.0-win-x64.zip",
            },
          ])
        );
      const res = await nodejsInstaller.getDownloadUrl(NpmMirror, "v22.14.0", "win-x64", ".zip");
      assert.isTrue(res.isOk());
      if (res.isOk()) {
        assert.equal(
          res.value,
          "https://cdn.npmmirror.com/binaries/node/v22.14.0/v22.14.0-win-x64.zip"
        );
      }
    });

    it("Official mirror - first fetchString fail", async () => {
      sandbox
        .stub(nodejsInstaller, "fetchString")
        .resolves(err(new InstallNodeJSError("test error")));
      const res = await nodejsInstaller.getDownloadUrl(
        OfficialMirror,
        "v22.14.0",
        "win-x64",
        ".zip"
      );
      assert.isTrue(res.isErr());
    });
    it("Official mirror - version not found", async () => {
      sandbox.stub(nodejsInstaller, "fetchString").resolves(ok(""));
      sandbox.stub(nodejsInstaller, "parseHtmlToGetUrl").returns(undefined);
      const res = await nodejsInstaller.getDownloadUrl(
        OfficialMirror,
        "v22.14.0",
        "win-x64",
        ".zip"
      );
      assert.isTrue(res.isErr());
      if (res.isErr()) {
        assert.include(
          res.error.message,
          getLocalizedString(
            "action.devTool.nodeInstaller.UnableToFind",
            "v22.14.0",
            "https://nodejs.org/dist/"
          )
        );
      }
    });
    it("Official mirror - second fetchString fail", async () => {
      sandbox
        .stub(nodejsInstaller, "fetchString")
        .onFirstCall()
        .resolves(ok(""))
        .onSecondCall()
        .resolves(err(new InstallNodeJSError("test error")));
      const res = await nodejsInstaller.getDownloadUrl(
        OfficialMirror,
        "v22.14.0",
        "win-x64",
        ".zip"
      );
      assert.isTrue(res.isErr());
    });
    it("Official mirror - package not found", async () => {
      sandbox
        .stub(nodejsInstaller, "fetchString")
        .onFirstCall()
        .resolves(ok(""))
        .onSecondCall()
        .resolves(ok(""));
      sandbox
        .stub(nodejsInstaller, "parseHtmlToGetUrl")
        .onFirstCall()
        .returns("https://nodejs.org/dist/v22.14.0/")
        .onSecondCall()
        .returns(undefined);
      const res = await nodejsInstaller.getDownloadUrl(
        OfficialMirror,
        "v22.14.0",
        "win-x64",
        ".zip"
      );
      assert.isTrue(res.isErr());
    });
    it("Official mirror - happy", async () => {
      sandbox
        .stub(nodejsInstaller, "fetchString")
        .onFirstCall()
        .resolves(ok(""))
        .onSecondCall()
        .resolves(ok(""));
      sandbox
        .stub(nodejsInstaller, "parseHtmlToGetUrl")
        .onFirstCall()
        .returns("https://nodejs.org/dist/v22.14.0/")
        .onSecondCall()
        .returns("https://nodejs.org/dist/v22.14.0/v22.14.0-win-x64.zip");
      const res = await nodejsInstaller.getDownloadUrl(
        OfficialMirror,
        "v22.14.0",
        "win-x64",
        ".zip"
      );
      assert.isTrue(res.isOk());
      if (res.isOk()) {
        assert.equal(res.value, "https://nodejs.org/dist/v22.14.0/v22.14.0-win-x64.zip");
      }
    });
  });

  describe("ensureNodeJS", () => {
    const context: any = {
      logProvider: new MockedLogProvider(),
      ui: new MockedUserInteraction(),
    };
    afterEach(() => {
      sandbox.restore();
    });
    it("system installed", async () => {
      sandbox.stub(NodeChecker, "getInstalledNodeVersion").resolves({ version: "v22.14.0" } as any);
      const res = await nodejsInstaller.ensureNodeJS(context, true, true);
      assert.isTrue(res.isOk());
      if (res.isOk()) {
        assert.deepEqual(res.value, { status: "ignore" });
      }
    });
    it("system not installed, user folder installed", async () => {
      sandbox.stub(NodeChecker, "getInstalledNodeVersion").resolves(null);
      sandbox.stub(nodejsInstaller, "getNameAndExt").returns({ name: "win-x64", ext: ".zip" });
      sandbox.stub(fs, "readdir").resolves(["node-v22.14.0-win-x64"] as any);
      const downloadDir = path.join(os.homedir(), `.${ConfigFolderName}`, "bin", "nodejs");
      const targetDir = path.join(downloadDir, "node-v22.14.0-win-x64");
      const res = await nodejsInstaller.ensureNodeJS(context, true, true);
      assert.isTrue(res.isOk());
      if (res.isOk()) {
        assert.deepEqual(res.value, { status: "ignore", installPath: targetDir });
      }
    });

    it("getBestMirror fail", async () => {
      sandbox.stub(NodeChecker, "getInstalledNodeVersion").resolves(null);
      sandbox.stub(nodejsInstaller, "getNameAndExt").returns({ name: "win-x64", ext: ".zip" });
      sandbox.stub(fs, "readdir").resolves([""] as any);
      sandbox.stub(nodejsInstaller, "getBestMirror").resolves(undefined);
      const res = await nodejsInstaller.ensureNodeJS(context, true, true);
      assert.isTrue(res.isErr());
      if (res.isErr()) {
        assert.include(
          res.error.message,
          getLocalizedString("action.devTool.nodeInstaller.NoMirror")
        );
      }
    });

    it("Confirm cancel", async () => {
      const NpmMirror: NodeDownloadMirror = {
        name: "NPM",
        url: "https://registry.npmmirror.com/-/binary/node/",
        indexJsonUrl: "https://cdn.npmmirror.com/binaries/node/index.json",
        packageUrl: "https://cdn.npmmirror.com/binaries/node/v22.14.0/v22.14.0-win-x64.zip",
      };
      sandbox.stub(NodeChecker, "getInstalledNodeVersion").resolves(null);
      sandbox.stub(nodejsInstaller, "getNameAndExt").returns({ name: "win-x64", ext: ".zip" });
      sandbox.stub(fs, "readdir").resolves([""] as any);
      sandbox.stub(nodejsInstaller, "getBestMirror").resolves(NpmMirror);
      sandbox.stub(context.ui, "confirm").resolves(err(new UserCancelError()));
      const res = await nodejsInstaller.ensureNodeJS(context, true, true);
      assert.isTrue(res.isErr());
      if (res.isErr()) {
        assert.isTrue(res.error instanceof UserCancelError);
      }
    });

    it("fetchBinary fail", async () => {
      const NpmMirror: NodeDownloadMirror = {
        name: "NPM",
        url: "https://registry.npmmirror.com/-/binary/node/",
        indexJsonUrl: "https://cdn.npmmirror.com/binaries/node/index.json",
        packageUrl: "https://cdn.npmmirror.com/binaries/node/v22.14.0/v22.14.0-win-x64.zip",
      };
      sandbox.stub(NodeChecker, "getInstalledNodeVersion").resolves(null);
      sandbox.stub(nodejsInstaller, "getNameAndExt").returns({ name: "win-x64", ext: ".zip" });
      sandbox.stub(fs, "readdir").resolves([""] as any);
      sandbox.stub(nodejsInstaller, "getBestMirror").resolves(NpmMirror);
      sandbox.stub(context.ui, "confirm").resolves(ok({ type: "success", result: true }));
      sandbox
        .stub(nodejsInstaller, "fetchBinary")
        .resolves(err(new InstallNodeJSError("test error")));
      const res = await nodejsInstaller.ensureNodeJS(context, true, true);
      assert.isTrue(res.isErr());
    });

    it("success", async () => {
      const NpmMirror: NodeDownloadMirror = {
        name: "NPM",
        url: "https://registry.npmmirror.com/-/binary/node/",
        indexJsonUrl: "https://cdn.npmmirror.com/binaries/node/index.json",
        packageUrl: "https://cdn.npmmirror.com/binaries/node/v22.14.0/v22.14.0-win-x64.zip",
        version: "v22.14.0",
      };
      sandbox.stub(NodeChecker, "getInstalledNodeVersion").resolves(null);
      sandbox.stub(nodejsInstaller, "getNameAndExt").returns({ name: "win-x64", ext: ".zip" });
      sandbox.stub(fs, "readdir").resolves([""] as any);
      sandbox.stub(nodejsInstaller, "getBestMirror").resolves(NpmMirror);
      sandbox.stub(context.ui, "confirm").resolves(ok({ type: "success", result: true }));
      sandbox.stub(nodejsInstaller, "fetchBinary").resolves(ok(Buffer.from("test buffer")));
      sandbox.stub(nodejsInstaller, "extractPackage").returns();
      const downloadDir = path.join(os.homedir(), `.${ConfigFolderName}`, "bin", "nodejs");
      const targetDir = path.join(downloadDir, "node-v22.14.0-win-x64");
      const res = await nodejsInstaller.ensureNodeJS(context, true, true);
      assert.isTrue(res.isOk());
      if (res.isOk()) {
        assert.equal(res.value.status, "installed");
        assert.equal(res.value.installPath, targetDir);
      }
    });
  });
});
