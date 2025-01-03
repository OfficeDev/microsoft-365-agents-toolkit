// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Helly Zhang <v-helzha@microsoft.com>
 */
import * as path from "path";
import * as fs from "fs-extra";
import { expect } from "chai";
import { describe } from "mocha";
import { EditorView, WebView, By, VSBrowser } from "vscode-extension-tester";
import {
  createNewProject,
  execCommandIfExist,
  execCommandIfExistFromTreeView,
} from "../../utils/vscodeOperation";
import { TreeViewTestContext, createSampleProject } from "./treeviewContext";
import {
  CommandPaletteCommands,
  Timeout,
  TreeViewCommands,
  TemplateProject,
  TemplateProjectFolder,
} from "../../utils/constants";
import { cleanUpLocalProject } from "../../utils/cleanHelper";
import { it } from "../../utils/it";
import { assert } from "chai";
import { join } from "path";
import { RetryHandler } from "../../utils/retryHandler";

describe("Create sample project and open sample view to download sample Tests", function () {
  this.timeout(Timeout.testCase);
  let treeViewTestContext: TreeViewTestContext;
  let testRootFolder: string;

  beforeEach(async function () {
    // ensure workbench is ready
    this.timeout(Timeout.prepareTestCase);
    treeViewTestContext = new TreeViewTestContext("samples");
    testRootFolder = treeViewTestContext.testRootFolder;
    await treeViewTestContext.before();
  });

  afterEach(async function () {
    this.timeout(Timeout.finishTestCase);
    // await treeViewTestContext.after();
  });

  it(
    "[auto] Sample app web view can ben open",
    {
      testPlanCaseId: 13022503,
      author: "v-helzha@microsoft.com",
    },
    async function () {
      try {
        //open sample view from command palette
        const driver = VSBrowser.instance.driver;
        await driver.sleep(Timeout.reloadWindow);

        await RetryHandler.retry(async () => {
          await execCommandIfExist("View: Toggle Full Screen");
        });

        {
          await new EditorView().closeAllEditors();
          console.log("Closed all opened editor view.");

          console.log("open new sample tab from tree view");
          await RetryHandler.retry(async () => {
            await execCommandIfExist(
              CommandPaletteCommands.SamplesCommand,
              Timeout.webView
            );
          });
          const webView = new WebView();
          console.log("find title");
          await driver.sleep(Timeout.webView);
          // wait for frame is ready
          await webView.switchToFrame();
          const element = await webView.findWebElement(
            By.xpath('//div[@class="sample-gallery"]//div[@class="title"]')
          );
          const text = await element.getText();
          await webView.switchBack();
          console.log("verify tab page");
          await driver.sleep(Timeout.webView);
          expect(text).has.string("Samples");
        }

        //open sample view from tree view
        {
          await new EditorView().closeAllEditors();
          await execCommandIfExistFromTreeView(
            TreeViewCommands.SamplesCommand,
            Timeout.webView
          );
          console.log("open new sample tab");
          await driver.sleep(Timeout.webView);
          const webView2 = new WebView();
          await webView2.switchToFrame();
          console.log("find title");
          await driver.sleep(Timeout.webView);
          const element = await webView2.findWebElement(
            By.xpath('//div[@class="sample-gallery"]//div[@class="title"]')
          );
          const text = await element.getText();
          await webView2.switchBack();
          console.log("verify tab page");
          await driver.sleep(Timeout.webView);
          expect(text).has.string("Samples");
          await new EditorView().closeAllEditors();
        }
      } catch (error) {
        console.log("error msg: ", error);
        await VSBrowser.instance.takeScreenshot("errorStep");
        assert.fail(error as string);
      }
    }
  );

  it(
    "[auto] Sample app can be downloaded successfully",
    {
      testPlanCaseId: 10241846,
      author: "v-helzha@microsoft.com",
    },
    async function () {
      const sample = "Tab App with Azure Backend";
      const samplePath = "hello-world-tab-with-backend";
      try {
        const sampleFolder = path.resolve(testRootFolder, `${samplePath}`);
        if (fs.existsSync(sampleFolder)) {
          await cleanUpLocalProject(sampleFolder);
        }
        await createSampleProject(sample);
        const filePath1 = path.join(
          testRootFolder,
          samplePath,
          "src",
          "index.tsx"
        );
        expect(fs.existsSync(filePath1), `${filePath1} must exist.`).to.eq(
          true
        );
        console.log("verify file successfully");
      } catch (error) {
        console.log("error msg: ", error);
        await VSBrowser.instance.takeScreenshot("errorStep");
        assert.fail(error as string);
      }
    }
  );
});
