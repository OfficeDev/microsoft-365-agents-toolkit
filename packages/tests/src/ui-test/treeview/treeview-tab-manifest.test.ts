// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Helly Zhang <v-helzha@microsoft.com>
 */
import {
  clearNotifications,
  createNewProject,
  getNotification,
} from "../../utils/vscodeOperation";
import { TreeViewTestContext, zipAppPackage } from "./treeviewContext";
import { createEnv } from "../remotedebug/remotedebugContext";
import { Timeout, Notification } from "../../utils/constants";
import { it } from "../../utils/it";

describe("Execute Build Teams Package", function () {
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
    "[auto] Execute Build Teams Package from Command Palette after Creating Tab Project",
    {
      testPlanCaseId: 14377966,
      author: "v-helzha@microsoft.com",
    },
    async function () {
      await createNewProject("tabnsso", treeViewTestContext.appName);
      await zipAppPackage("dev");
      await getNotification(
        Notification.UnresolvedPlaceholderError,
        Timeout.shortTimeWait
      );

      await clearNotifications();
      await zipAppPackage("local");
      await getNotification(
        Notification.UnresolvedPlaceholderError,
        Timeout.shortTimeWait
      );

      await clearNotifications();
      await createEnv("staging");
      await zipAppPackage("staging");
      await getNotification(
        Notification.UnresolvedPlaceholderError,
        Timeout.shortTimeWait
      );
    }
  );
});
