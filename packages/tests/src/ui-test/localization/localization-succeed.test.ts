// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { Notification, Timeout } from "../../utils/constants";
import {
  getNotification,
  openExistingProject,
} from "../../utils/vscodeOperation";
import { it } from "../../utils/it";
import * as path from "path";
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
});
