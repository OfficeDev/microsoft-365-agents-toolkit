// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Helly Zhang <v-helzha@microsoft.com>
 */
import {
  By,
  EditorView,
  InputBox,
  VSBrowser,
  WebView,
} from "vscode-extension-tester";
import { expect } from "chai";
import { execCommandIfExist } from "../../utils/vscodeOperation";
import { TreeViewTestContext } from "./treeviewContext";
import {
  CommandPaletteCommands,
  CreateProjectQuestion,
  Timeout,
} from "../../utils/constants";
import { delay, RetryHandler } from "../../utils/retryHandler";
import { it } from "../../utils/it";

describe("Openning Quick Start Tests", function () {
  this.timeout(Timeout.testCase);
  let treeViewTestContext: TreeViewTestContext;

  beforeEach(async function () {
    // ensure workbench is ready
    this.timeout(Timeout.prepareTestCase);
    treeViewTestContext = new TreeViewTestContext("treeview");
    await treeViewTestContext.before();
  });

  afterEach(async function () {
    this.timeout(Timeout.finishTestCase);
    await treeViewTestContext.after();
  });

  it(
    "[auto] [QuickStart] Check contents",
    {
      testPlanCaseId: 12933026,
      author: "v-helzha@microsoft.com",
    },
    async function () {
      const driver = VSBrowser.instance.driver;
      await driver.sleep(Timeout.reloadWindow);

      // get started page for "Build a Notification Bot"
      await RetryHandler.retry(async () => {
        await execCommandIfExist("View: Toggle Full Screen");
      });

      await new EditorView().closeAllEditors();
      console.log("Closed all opened editor view.");

      await execCommandIfExist(
        CommandPaletteCommands.QuickStartCommand,
        Timeout.webView
      );
      const input = await InputBox.create();
      await input.selectQuickPick(CreateProjectQuestion.BuildNotificationBot);

      const webView = new WebView();

      const element = await webView.findWebElement(
        By.className("category-description-container")
      );
      const type1Title = await element.getText();
      expect(type1Title).has.string(CreateProjectQuestion.BuildNotificationBot);

      // Check item "Get your environment ready"
      const type1Item1 = await getExpandedButton(
        webView,
        false,
        "Get your environment ready"
      );
      const type1Item1Button = await type1Item1?.findElement(
        By.css(".button-container .monaco-button")
      );
      const type1Item1ButtonValue = await type1Item1Button.getText();
      expect(type1Item1ButtonValue).has.string("Check Prerequisites");
      console.log('Found the button "Check Prerequisites"');

      // Check item "Create a notification bot"
      const type1Item2 = await getExpandedButton(
        webView,
        false,
        "Build a notification bot"
      );
      const type1Item2Button = await type1Item2?.findElement(
        By.css(".button-container .monaco-button")
      );
      const type1Item2ButtonValue = await type1Item2Button.getText();
      expect(type1Item2ButtonValue).has.string("Build Notification Bot");
      console.log('Found the button "Build Notification Bot"');

      // get started page for "Build a Declarative Agent"
      await execCommandIfExist(
        CommandPaletteCommands.QuickStartCommand,
        Timeout.webView
      );
      const input2 = await InputBox.create();
      await input2.selectQuickPick(CreateProjectQuestion.BuildDeclarativeAgent);

      const webView2 = new WebView();

      const element2 = await webView2.findWebElement(
        By.className("category-description-container")
      );
      const type2Title = await element2.getText();
      expect(type2Title).has.string(
        CreateProjectQuestion.BuildDeclarativeAgent
      );

      // Check item "Get your environment ready"
      const type2Item1 = await getExpandedButton(
        webView,
        false,
        "Get your environment ready"
      );
      const type2Item1Button = await type2Item1?.findElement(
        By.css(".button-container .monaco-button")
      );
      const type2Item1ButtonValue = await type2Item1Button.getText();
      expect(type2Item1ButtonValue).has.string("Check Copilot License");
      console.log('Found the button "Check Copilot License"');

      // Check item "Build a declarative agent"
      const type2Item2 = await getExpandedButton(
        webView,
        false,
        "Build a declarative agent"
      );
      const type2Item2Button = await type2Item2?.findElement(
        By.css(".button-container .monaco-button")
      );
      const type2Item2ButtonValue = await type2Item2Button.getText();
      expect(type2Item2ButtonValue).has.string("Build Declarative Agent");
      console.log('Found the button "Build Declarative Agent"');
    }
  );
});

async function getExpandedButton(
  webView: WebView,
  expended = true,
  content = "Build your first app"
) {
  if (!expended) {
    const collapsedButtons = await webView.findWebElements(
      By.xpath('//button[@class="getting-started-step"]')
    );
    await delay(Timeout.shortTimeWait);
    for (const button of collapsedButtons) {
      const item = await button.findElement(By.css("h3"));
      const itemContext = await item.getText();
      if (itemContext.includes(content)) {
        await button.click();
        await delay(Timeout.shortTimeWait);
        break;
      }
    }
  }
  const button = await webView.findWebElement(
    By.xpath('//button[@class="getting-started-step expanded"]')
  );
  return button;
}
