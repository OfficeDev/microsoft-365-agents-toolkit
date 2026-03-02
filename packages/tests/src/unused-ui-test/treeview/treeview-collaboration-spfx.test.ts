// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Helly Zhang <v-helzha@microsoft.com>
 */
import { expect } from "chai";
import {
  CommandPaletteCommands,
  Timeout,
  Notification,
} from "../../utils/constants";
import {
  createNewProject,
  execCommandIfExist,
  getNotification,
} from "../../utils/vscodeOperation";
import path = require("path");
import { VSBrowser } from "vscode-extension-tester";
import { Env } from "../../utils/env";
import {
  addCollaborators,
  getAllCollaborators,
} from "../../utils/collaborationUtils";
import { it } from "../../utils/it";
import { TreeViewTestContext } from "./treeviewContext";

describe("Collaborator Tests SPFX", function () {
  this.timeout(Timeout.testCase);
  let remoteDebugTestContext: TreeViewTestContext;
  let testRootFolder: string;
  let appName: string;
  const appNameCopySuffix = "copy";
  let newAppFolderName: string;
  let projectPath: string;

  beforeEach(async function () {
    // ensure workbench is ready
    this.timeout(Timeout.prepareTestCase);
    remoteDebugTestContext = new TreeViewTestContext("spfx");
    testRootFolder = remoteDebugTestContext.testRootFolder;
    appName = remoteDebugTestContext.appName;
    newAppFolderName = appName + appNameCopySuffix;
    projectPath = path.resolve(testRootFolder, newAppFolderName);
    await remoteDebugTestContext.before();
  });

  afterEach(async function () {
    this.timeout(Timeout.finishTestCase);
  });

  it(
    "[auto] Add collaborator for spfx project",
    {
      testPlanCaseId: 12933031,
      author: "v-helzha@microsoft.com",
    },
    async function () {
      const creator = Env.username;
      const collaborator = Env.collaborator;

      // //create SPFx project
      const driver = VSBrowser.instance.driver;
      await createNewProject("spfx", appName);
      console.log("Finish create SPFX project");

      await execCommandIfExist(CommandPaletteCommands.ProvisionCommand);
      await driver.sleep(Timeout.spfxProvision);
      await getNotification(
        Notification.ProvisionSucceeded,
        Timeout.shortTimeWait
      );
      console.log("Finish provision SPFX project");

      {
        const findCollaborator = await getAllCollaborators();
        expect(findCollaborator.includes(creator as string)).to.be.true;
      }

      await addCollaborators(collaborator);

      {
        const findCollaborator = await getAllCollaborators();
        expect(
          findCollaborator.includes((collaborator as string)?.split("@")[0])
        ).to.be.true;
        expect(findCollaborator.includes((creator as string)?.split("@")[0])).to
          .be.true;
      }
    }
  );
});
