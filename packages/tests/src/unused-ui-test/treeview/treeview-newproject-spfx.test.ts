// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Helly Zhang <v-helzha@microsoft.com>
 */
import * as path from "path";
import * as fs from "fs-extra";
import { expect } from "chai";
import { Timeout } from "../../utils/constants";
import { TreeViewTestContext } from "./treeviewContext";
import { createNewProject } from "../../utils/vscodeOperation";
import { it } from "../../utils/it";

describe("New project Tests", function () {
  this.timeout(Timeout.testCase);
  let treeViewTestContext: TreeViewTestContext;
  let testRootFolder: string;
  const appNameCopySuffix = "copy";
  let newAppFolderName: string;
  let projectPath: string;

  beforeEach(async function () {
    // ensure workbench is ready
    this.timeout(Timeout.prepareTestCase);
    treeViewTestContext = new TreeViewTestContext("treeview");
    testRootFolder = treeViewTestContext.testRootFolder;
    await treeViewTestContext.before();
  });

  afterEach(async function () {
    this.timeout(Timeout.finishTestCase);
    await treeViewTestContext.after();
  });

  it(
    "[auto] Create SPFx project",
    {
      testPlanCaseId: 11967093,
      author: "v-helzha@microsoft.com",
    },
    async function () {
      const appName = treeViewTestContext.appName;
      await createNewProject("spfx", appName);
      newAppFolderName = appName + appNameCopySuffix;
      projectPath = path.resolve(testRootFolder, newAppFolderName);
      const filePath = path.join(projectPath, "src", "src", "index.ts");
      expect(fs.existsSync(filePath), `${filePath} must exist.`).to.eq(true);
    }
  );
});
