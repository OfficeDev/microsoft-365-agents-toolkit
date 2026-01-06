// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import * as os from "os";
import * as path from "path";
import fs from "fs-extra";

import { ConfigFolderName } from "@microsoft/teamsfx-api";
import sinon from "sinon";

import { ServiceLogWriter } from "../../../../src/cmds/preview/serviceLogWriter";
import { expect } from "../../utils";

describe("ServiceLogWriter", () => {
  const cliLogFolderName = "cli-log";
  const localPreviewLogFolderName = "local-preview";
  const localPreviewLogFolder = path.join(
    os.homedir(),
    `.${ConfigFolderName}`,
    cliLogFolderName,
    localPreviewLogFolderName
  );

  const sandbox = sinon.createSandbox();
  const folders = new Set<string>();
  const logs = new Map<string, string>();

  beforeEach(() => {
    sandbox.stub(fs, "ensureDir").callsFake(async (dir: string) => {
      const basename = path.basename(dir);
      if (path.join(localPreviewLogFolder, basename) === dir) {
        folders.add(basename);
      }
    });
    sandbox.stub(fs, "readdir").callsFake(async (dir: string | Buffer) => {
      if (dir === localPreviewLogFolder) {
        return Array.from(folders).map((dir) => {
          return path.dirname(dir);
        });
      }
      return [];
    });
    sandbox.stub(fs, "remove").callsFake(async (dir: string) => {
      const basename = path.basename(dir);
      if (path.join(localPreviewLogFolder, basename) === dir) {
        folders.delete(basename);
      }
    });
    sandbox.stub(fs, "ensureFile").callsFake(async (file: string) => {
      if (!logs.has(file)) {
        logs.set(file, "");
      }
    });
    sandbox.stub(fs, "appendFile").callsFake(async (file: string | number | Buffer, data: any) => {
      logs.set(file as string, logs.get(file as string) + data);
    });
    sandbox.stub(fs, "pathExists").callsFake(async (file: string) => {
      return logs.has(file);
    });
  });

  afterEach(() => {
    sandbox.restore();
  });

  describe("init", () => {
    it("happy path", async () => {
      const datetime = new Date().toISOString();
      const stub = sandbox.stub(Date.prototype, "toISOString");
      stub.callsFake(() => {
        return datetime;
      });
      const serviceLogWriter = new ServiceLogWriter();
      await serviceLogWriter.init();
      expect(folders.size).equals(1);
      expect(folders.entries().next().value![0]).equals(
        datetime.replace(/:/g, "_").replace(/\./g, "_")
      );
      stub.restore();
    });
  });

  describe("write and getLogFile", async () => {
    it("happy path", async () => {
      const datetime = new Date().toISOString();
      const stub = sandbox.stub(Date.prototype, "toISOString");
      stub.callsFake(() => {
        return datetime;
      });
      const serviceLogWriter = new ServiceLogWriter();
      await serviceLogWriter.init();
      const serviceTitle = "test start";
      const message = "test started successfully.";
      await serviceLogWriter.write(serviceTitle, message);
      const logFile = await serviceLogWriter.getLogFile(serviceTitle);
      expect(logFile).to.not.equal(undefined);
      expect(logFile).equals(
        path.join(
          localPreviewLogFolder,
          datetime.replace(/:/g, "_").replace(/\./g, "_"),
          `${serviceTitle.split(" ").join("-")}.log`
        )
      );
      expect(logs.has(logFile as string)).equals(true);
      expect(logs.get(logFile as string)).equals(message);
    });
  });
});
