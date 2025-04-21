// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Helly Zhang <v-helzha@microsoft.com>
 */
import * as path from "path";
import * as fs from "fs-extra";
import { expect } from "chai";
import { Timeout, Lang } from "../../utils/constants";
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
    "[auto] Create Tab typescript project and validation",
    {
      testPlanCaseId: 10791728,
      author: "v-helzha@microsoft.com",
    },
    async function () {
      const appName = treeViewTestContext.appName;
      await createNewProject("tab", appName, { lang: Lang.TS });
      newAppFolderName = appName + appNameCopySuffix;
      projectPath = path.resolve(testRootFolder, newAppFolderName);
      const filePath1 = path.join(projectPath, "src", "index.tsx");
      expect(fs.existsSync(filePath1), `${filePath1} must exist.`).to.eq(true);
    }
  );
});
