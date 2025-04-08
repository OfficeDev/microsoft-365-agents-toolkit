// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import {
  ActivityBar,
  BottomBarPanel,
  By,
  InputBox,
  SideBarView,
  VSBrowser,
  WebElement,
} from "vscode-extension-tester";
import { Extension, Timeout, TreeViewCommands } from "./constants";
import { clearNotifications, openTerminalView } from "./vscodeOperation";
import { Executor } from "../utils/executor";

const listCollaborator = "List Microsoft 365 Teams App (with AAD App) Owners";
const grantPermission = "Manage M365 Teams App (with AAD app) Collaborators";

export async function getAllCollaboratorsCLI(projectPath: string) {
  const { stdout, stderr } = await Executor.listAppOwners(projectPath);
  if (stderr) {
    throw new Error(stderr);
  }
  return stdout;
}

export async function addCollaboratorCLI(
  projectPath: string,
  email: string,
  teamsManifestFilePath: string
): Promise<void> {
  const { stdout, stderr } = await Executor.addAppOwner(
    projectPath,
    email,
    teamsManifestFilePath
  );
  if (stderr) {
    throw new Error(stderr);
  }
  console.log(stdout);
}

export async function getAllCollaborators(): Promise<string> {
  const driver = VSBrowser.instance.driver;
  await clearNotifications();
  console.log("openTerminalView");
  await openTerminalView();
  console.log("openOutputView");
  const pannel = new BottomBarPanel();
  const output = await pannel.openOutputView();
  console.log("Teams Toolkit");
  try {
    const maximize = await pannel.findElement(
      By.css("a.action-label.codicon.codicon-panel-maximize")
    );
    await maximize.click();
    await driver.sleep(Timeout.shortTimeWait);
  } catch {
    console.log("already maximized");
  }
  // This api is not work on macos, it will throw: Error: Channel Teams Toolkit not found
  await output.selectChannel("Microsoft 365 Agents");
  // Clear output
  console.log("start Clear output");
  await output.clearText();
  console.log("Clear output");

  // Trigger list collaborators
  const extensionView = await openExtensionView();
  const sections = await getSections(extensionView);
  await runManageCollaborators(sections);
  const addCollaboratorInput = await InputBox.create();
  await addCollaboratorInput.selectQuickPick("List App Owners");
  await driver.sleep(Timeout.input);
  const teamsApp = await addCollaboratorInput.findElement(
    By.xpath(".//div[@role='checkbox' and @data-index='1']")
  );
  await teamsApp.click();
  await driver.sleep(Timeout.input);
  await addCollaboratorInput.confirm();
  await driver.sleep(Timeout.input);
  await addCollaboratorInput.selectQuickPick("manifest.json");
  await driver.sleep(Timeout.input);
  await driver.sleep(Timeout.shortTimeWait);
  console.log("Click list collaborators");

  // Get output
  console.log("Get output");
  const text = await output.getText();

  console.log("Output: " + text);
  return text;
}

export async function openExtensionView(
  extensionName = Extension.displayName
): Promise<SideBarView> {
  const activityBar = new ActivityBar();
  const view = await activityBar.getViewControl(extensionName);
  if (!view) {
    throw "No view opened";
  }
  const extensionView = await view.openView();
  return extensionView;
}

export async function getSections(
  extensionView: SideBarView
): Promise<WebElement[]> {
  // if extensionView is not fully displayed, getContent() will fail
  // const extensionContent = extensionView.getContent();
  // const sections = await extensionContent.getSections();

  const sections = await extensionView.findElements(
    By.className("split-view-view")
  );
  return sections;
}

export async function runManageCollaborators(
  sections: WebElement[]
): Promise<void> {
  const driver = VSBrowser.instance.driver;
  for (const section of sections) {
    const sectionTitle = await section.findElement(By.css(".title")).getText();
    if (sectionTitle.toUpperCase() === TreeViewCommands.EnvSectionName) {
      const treeItems = await section.findElements(By.css(".monaco-list-row"));
      for (const treeItem of treeItems) {
        const treeItemText = await treeItem.getText();
        if (treeItemText === "dev(Provisioned)") {
          console.log("Found environment: " + treeItemText);

          // show action items
          await treeItem.click();
          await driver.sleep(Timeout.shortTimeWait);

          const actionItem = await treeItem.findElement(
            By.css("a.action-label.codicon.codicon-organization")
          );
          await actionItem.click();
          break;
        }
      }
      break;
    }
  }
}

export async function addCollaborators(collaborator: string): Promise<void> {
  await clearNotifications();
  const extensionView = await openExtensionView();
  const sections = await getSections(extensionView);
  await runManageCollaborators(sections);
  const driver = VSBrowser.instance.driver;
  const addCollaboratorInput = await InputBox.create();
  await addCollaboratorInput.selectQuickPick("Add App Owners");
  await driver.sleep(Timeout.input);
  const teamsApp = await addCollaboratorInput.findElement(
    By.xpath(".//div[@role='checkbox' and @data-index='1']")
  );
  await teamsApp.click();
  await driver.sleep(Timeout.input);
  await addCollaboratorInput.confirm();
  await driver.sleep(Timeout.input);
  await addCollaboratorInput.selectQuickPick("manifest.json");
  await driver.sleep(Timeout.input);
  await addCollaboratorInput.setText(collaborator as string);
  console.log("Input collaborator: " + collaborator);
  await driver.sleep(Timeout.shortTimeWait);
  await addCollaboratorInput.confirm();
  await driver.sleep(Timeout.addCollaborator);
}
