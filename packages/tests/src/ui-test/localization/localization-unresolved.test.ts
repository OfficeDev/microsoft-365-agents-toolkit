// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
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
import * as fs from "fs-extra";
import { validateFileExist } from "../../utils/commonUtils";
import { validateAppPackage, zipAppPackage } from "../treeview/treeviewContext";

describe("Env support for localization.json", function () {
  this.timeout(Timeout.testCase);

  beforeEach(async function () {
    this.timeout(Timeout.prepareTestCase);
  });

  afterEach(async function () {
    this.timeout(Timeout.finishTestCase);
  });

  it(
    "Check Validate app package fails when localization file has unresolved envs",
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
      console.log("Run Zip App Package");
      await zipAppPackage("dev");
      await getNotification(
        Notification.ZipAppPackageSucceeded,
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
      await zipAppPackage("dev");
      await getNotification(
        Notification.UnresolvedPlaceholderError,
        Timeout.shortTimeWait
      );
    }
  );
});
