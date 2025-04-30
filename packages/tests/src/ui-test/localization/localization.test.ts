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
import { dotenvUtil } from "../../utils/envUtil";
import { validateAppPackage, zipAppPackage } from "../treeview/treeviewContext";
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
    "Env support for localization.json - Validate app package succeed",
    {
      testPlanCaseId: 30481064,
      author: "v-helzha@microsoft.com",
    },
    async function () {
      const resourceFolder = path.resolve(
        __dirname,
        "../../../src/ui-test/case-resources/"
      );
      const projectPath = path.resolve(
        resourceFolder,
        "localization-agent-succeed/agent"
      );
      console.log("Project path: ", projectPath);
      await openExistingProject(projectPath);
      console.log("Run Zip App Package");
      await zipAppPackage("dev");
      await getNotification(
        Notification.ZipAppPackageSucceeded,
        Timeout.shortTimeWait
      );
      console.log("Validate app manifest schema");
      await validateAppPackage("dev");
      await getNotification(
        Notification.appManifestSchemaSucceeded,
        Timeout.shortTimeWait
      );
    }
  );

  it(
    "Check Validate app manifest fails when localization file has unresolved envs",
    {
      testPlanCaseId: 30481072,
      author: "v-annefu@microsoft.com",
    },
    async function () {
      const resourceFolder = path.resolve(
        __dirname,
        "../../../src/ui-test/case-resources/"
      );
      const projectPath = path.resolve(
        resourceFolder,
        "localization-agent-unresolved/agent"
      );
      console.log("Project path: ", projectPath);
      await openExistingProject(projectPath);
      console.log("Validate app manifest schema");
      await validateAppPackage("dev");
      await getNotification(
        Notification.appManifestSchemaSucceeded,
        Timeout.shortTimeWait
      );

      //replace en.json
      const enjsonFile = path.resolve(projectPath, "appPackage/loc", "en.json");
      const unResolvedFile = path.resolve(resourceFolder, "en.json");
      if (await fs.pathExists(enjsonFile)) {
        await fs.copy(unResolvedFile, enjsonFile, {
          overwrite: true,
        });
      }
      await validateAppPackage("dev");
      await getNotification(
        Notification.UnresolvedPlaceholderError,
        Timeout.shortTimeWait
      );
    }
  );

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

  it(
    "Check Different level localization files are generated correctly when creating app package",
    {
      testPlanCaseId: 30481031,
      author: "v-ivanchen@microsoft.com",
    },

    async function () {
      const resourceFolder = path.resolve(
        __dirname,
        "../../../src/ui-test/case-resources/"
      );
      const projectPath = path.resolve(
        resourceFolder,
        "localization-agent-different-level/agent"
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
      const enFile = zipEntries.find((x) => x.entryName.includes("en.json"));
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
