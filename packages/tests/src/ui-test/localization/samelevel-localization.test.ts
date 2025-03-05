// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Anne Fu <v-annefu@microsoft.com>
 */
import { Notification, Timeout } from "../../utils/constants";
import {
  getNotification,
  openExistingProject,
} from "../../utils/vscodeOperation";
import { it } from "../../utils/it";
import * as path from "path";
import AdmZip from "adm-zip";
import { assert, expect } from "chai";
import { zipAppPackage } from "../treeview/treeviewContext";
import { dotenvUtil } from "../../utils/envUtil";
import * as fs from "fs-extra";

describe("Env support for localization.json", function () {
  this.timeout(Timeout.testCase);

  beforeEach(async function () {
    this.timeout(Timeout.prepareTestCase);
  });

  afterEach(async function () {
    this.timeout(Timeout.finishTestCase);
  });

  it(
    "Check Same level localization files are generated correctly when creating app package",
    {
      testPlanCaseId: 30481003,
      author: "v-annefu@microsoft.com",
    },

    async function () {
      const resourceFolder = path.resolve(
        __dirname,
        "../../../src/ui-test/case-resources/"
      );
      const projectPath = path.resolve(
        resourceFolder,
        "localization-agent-same-level/agent"
      );
      console.log("open project path" + projectPath);
      await openExistingProject(projectPath);
      console.log("Run Zip App Package");
      await zipAppPackage("dev");
      await getNotification(
        Notification.ZipAppPackageSucceeded,
        Timeout.shortTimeWait
      );
      const appPackageZipPath = path.join(
        projectPath,
        "appPackage",
        "build",
        "appPackage.dev.zip"
      );
      const zip = new AdmZip(appPackageZipPath);
      const zipEntries = zip.getEntries();
      const enFile = zipEntries.find((x) => x.entryName === "en.json");
      console.log("verify file successfully");
      if (!enFile) {
        assert.fail("en file not found error");
      }
      const enfileString = enFile.getData().toString();
      const envPath = path.resolve(projectPath, "env", ".env.dev");
      const config = dotenvUtil.deserialize(fs.readFileSync(envPath, "utf-8"));
      const teamsappAppversion = config.obj["TEAMS_APP_VERSION"];
      const envOwner = config.obj["OWNER"];
      console.log("env teamsappversion:", teamsappAppversion);
      console.log("env owner:", envOwner);
      expect(enfileString.includes(teamsappAppversion)).to.be.true;
      expect(enfileString.includes(envOwner)).to.be.true;
    }
  );
});
