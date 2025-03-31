// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Ivan Chen <v-ivanchen@microsoft.com>
 */
import * as path from "path";
import * as fs from "fs-extra";
import * as os from "os";
import { expect, assert } from "chai";
import {
  Timeout,
  TreeViewCommands,
  CreateProjectQuestion,
} from "../../utils/constants";
import { TreeViewTestContext } from "./treeviewContext";
import {
  createNewProject,
  execCommandIfExistFromTreeView,
  inputFolderPath,
  openExistingProject,
} from "../../utils/vscodeOperation";
import { InputBox, VSBrowser } from "vscode-extension-tester";
import { it } from "../../utils/it";

describe("New project in existing project Tests", function () {
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
      testPlanCaseId: "",
      author: "v-ivanchen@microsoft.com",
    },
    async function () {
      const appName = treeViewTestContext.appName;
      await createNewProject("tab", appName, { lang: "TypeScript" });
      newAppFolderName = appName + appNameCopySuffix;
      projectPath = path.resolve(testRootFolder, newAppFolderName);
      const filePath1 = path.join(projectPath, "src", "index.tsx");
      expect(fs.existsSync(filePath1), `${filePath1} must exist.`).to.eq(true);

      // create new project in existing project
      const newAppName = appName + appNameCopySuffix;
      console.log("create new project in existing project");
      const driver = VSBrowser.instance.driver;
      await execCommandIfExistFromTreeView(
        TreeViewCommands.CreateProjectCommand,
        Timeout.webView
      );

      const input = await InputBox.create();
      await input.selectQuickPick(CreateProjectQuestion.Tab);
      await driver.sleep(Timeout.input);
      await input.selectQuickPick("Basic Tab");
      await driver.sleep(Timeout.input);

      // Choose programming language
      await input.selectQuickPick("TypeScript");
      await driver.sleep(Timeout.input);

      // Input folder path
      console.log("choose project path: ", testRootFolder);
      await input.selectQuickPick("Browse...");
      await inputFolderPath(driver, input, testRootFolder);
      await driver.sleep(Timeout.input);
      if (os.type() === "Windows_NT") {
        await input.sendKeys("\\");
      } else if (os.type() === "Linux") {
        await input.sendKeys("/");
      }
      await input.confirm();

      // Input App Name
      if ((await input.getTitle()) === "Application Name") {
        await input.setText(newAppName);
        await driver.sleep(Timeout.input);
        await input.confirm();
      } else {
        const title = await input.getTitle();
        console.log("title:", title);
        assert.fail("Failed to input app name");
      }

      await driver.sleep(Timeout.input);

      const newProjectPath = path.resolve(testRootFolder, newAppName);
      const newProjectCopyPath = path.resolve(
        testRootFolder,
        newAppName + appNameCopySuffix
      );
      console.log("copy path: ", projectPath, " to: ", newProjectPath);
      await fs.mkdir(newProjectCopyPath);
      const filterFunc = (src: string) =>
        src.indexOf("node_modules") > -1 ? false : true;
      await fs.copy(projectPath, newProjectCopyPath, { filter: filterFunc });
      console.log("open project path");
      await openExistingProject(newProjectCopyPath);
    }
  );
});
