// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Helly Zhang <v-helzha@microsoft.com>
 */
import { expect } from "chai";
import { TreeViewTestContext, checkSectionContent } from "./treeviewContext";
import { Timeout, TreeViewCommands } from "../../utils/constants";
import {
  createNewProject,
  openExistingProject,
} from "../../utils/vscodeOperation";
import { it } from "../../utils/it";
import * as path from "path";
import * as fs from "fs-extra";

describe("Check command name in command palette and tree view content Tests", function () {
  this.timeout(Timeout.testCase);
  let treeViewTestContext: TreeViewTestContext;
  let testRootFolder: string;
  let appName: string;

  beforeEach(async function () {
    // ensure workbench is ready
    this.timeout(Timeout.prepareTestCase);
    treeViewTestContext = new TreeViewTestContext("treeview");
    testRootFolder = treeViewTestContext.testRootFolder;
    appName = treeViewTestContext.appName;
    await treeViewTestContext.before();
  });

  afterEach(async function () {
    this.timeout(Timeout.finishTestCase);
    await treeViewTestContext.after();
  });

  it(
    "[auto] All treeview items display correctly",
    {
      testPlanCaseId: 10241837,
      author: "v-helzha@microsoft.com",
    },
    async function () {
      await createNewProject("tabnsso", appName);
      const includeDevelopmentSection = await checkSectionContent(
        TreeViewCommands.DevelopmentSectionName,
        TreeViewCommands.DevelopmentSectionItems
      );
      expect(
        includeDevelopmentSection,
        `${TreeViewCommands.DevelopmentSectionName} does not show all elements.`
      ).equal(true);
    }
  );

  it(
    "[auto] Check office dev add-in treeview items display correctly",
    {
      testPlanCaseId: 27569380,
      author: "xuruiyao@microsoft.com",
    },
    async function () {
      const projectPath = path.resolve(
        testRootFolder,
        "../",
        "src",
        "ui-test",
        "treeview",
        "word-xml-addin"
      );
      const projectCopyPath = path.resolve(testRootFolder, appName + "copy");
      console.log("copy path: ", projectPath, " to: ", projectCopyPath);
      await fs.mkdir(projectCopyPath);
      const filterFunc = (src: string) =>
        src.indexOf("node_modules") > -1 ? false : true;
      await fs.copy(projectPath, projectCopyPath, { filter: filterFunc });
      console.log("open project path");
      await openExistingProject(projectCopyPath);

      const includeDevelopmentSection = await checkSectionContent(
        TreeViewCommands.OfficeDevDevelopmentSectionName,
        TreeViewCommands.OfficeDevDevelopmentSectionItems
      );

      const includeLifeCycleSection = await checkSectionContent(
        TreeViewCommands.OfficeDevLifeCycleSectionName,
        TreeViewCommands.OfficeDevLifeCycleSectionItems
      );

      const includeUtilitySection = await checkSectionContent(
        TreeViewCommands.OfficeDevUtilitySectionName,
        TreeViewCommands.OfficeDevUtilitySectionItems
      );

      const includeHelpAndFeedbackSection = await checkSectionContent(
        TreeViewCommands.OfficeDevHelpAndFeedBackSectionName,
        TreeViewCommands.OfficeDevHelpAndFeedBackSectionItems
      );

      expect(
        includeDevelopmentSection,
        `${TreeViewCommands.OfficeDevDevelopmentSectionName} does not show all elements.`
      ).equal(true);

      expect(
        includeLifeCycleSection,
        `${TreeViewCommands.OfficeDevLifeCycleSectionName} does not show all elements.`
      ).equal(true);

      expect(
        includeUtilitySection,
        `${TreeViewCommands.OfficeDevUtilitySectionName} does not show all elements.`
      ).equal(true);

      expect(
        includeHelpAndFeedbackSection,
        `${TreeViewCommands.OfficeDevHelpAndFeedBackSectionName} does not show all elements.`
      ).equal(true);
    }
  );
});
