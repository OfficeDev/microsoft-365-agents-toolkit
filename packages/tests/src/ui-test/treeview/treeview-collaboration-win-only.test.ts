// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Helly Zhang <v-helzha@microsoft.com>
 */
import * as fs from "fs-extra";
import { expect } from "chai";
import { Timeout } from "../../utils/constants";
import {
  createNewProject,
  execCommandIfExist,
  ensureExtensionActivated,
} from "../../utils/vscodeOperation";
import {
  RemoteDebugTestContext,
  runProvision,
  provisionProject,
} from "../remotedebug/remotedebugContext";
import path = require("path");
import { VSBrowser } from "vscode-extension-tester";
import { Env } from "../../utils/env";
import {
  getAllCollaboratorsCLI,
  addCollaboratorCLI,
} from "../../utils/collaborationUtils";
import { it } from "../../utils/it";

describe("Collaborator Tests", function () {
  this.timeout(Timeout.testCase);
  let remoteDebugTestContext: RemoteDebugTestContext;
  let testRootFolder: string;
  let appName: string;
  const appNameCopySuffix = "copy";
  let newAppFolderName: string;
  let projectPath: string;

  beforeEach(async function () {
    // ensure workbench is ready
    this.timeout(Timeout.prepareTestCase);
    remoteDebugTestContext = new RemoteDebugTestContext("tab");
    testRootFolder = remoteDebugTestContext.testRootFolder;
    appName = remoteDebugTestContext.appName;
    newAppFolderName = appName + appNameCopySuffix;
    projectPath = path.resolve(testRootFolder, newAppFolderName);
    // await remoteDebugTestContext.before();
    await fs.ensureDir(testRootFolder);
    await ensureExtensionActivated();
  });

  afterEach(async function () {
    this.timeout(Timeout.finishTestCase);
    // await remoteDebugTestContext.after();

    // Close the folder and cleanup local sample project
    await execCommandIfExist("Workspaces: Close Workspace", Timeout.webView);
    console.log(`[Successfully] start to clean up for ${projectPath}`);
    await remoteDebugTestContext.cleanUp(
      appName,
      projectPath,
      true,
      false,
      false
    );
  });

  it(
    "[auto] Add collaborator",
    {
      testPlanCaseId: 11966674,
      author: "v-helzha@microsoft.com",
    },
    async function () {
      const creator = Env.username;
      const collaborator = Env.collaborator;
      //create tab project
      const driver = VSBrowser.instance.driver;
      await createNewProject("tabnsso", appName);
      await provisionProject(appName, projectPath);

      {
        const findCollaborator = await getAllCollaboratorsCLI(projectPath);
        console.log(findCollaborator);
        expect(findCollaborator.includes(creator as string)).to.be.true;
      }

      const teamsManifestFilePath = path.resolve(
        projectPath,
        "appPackage",
        "manifest.json"
      );
      await addCollaboratorCLI(
        projectPath,
        collaborator,
        teamsManifestFilePath
      );
      // cli not support
      // {
      //   const findCollaborator = await getAllCollaboratorsCLI(projectPath);
      //   console.log(findCollaborator);
      //   expect(
      //     findCollaborator.includes((collaborator as string)?.split("@")[0])
      //   ).to.be.true;
      //   expect(findCollaborator.includes((creator as string)?.split("@")[0])).to
      //     .be.true;
      // }
    }
  );
});
