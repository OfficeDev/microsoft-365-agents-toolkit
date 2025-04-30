// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { hooks } from "@feathersjs/hooks/lib";
import {
  ConfigFolderName,
  CoreCallbackEvent,
  FxError,
  Inputs,
  ok,
  Platform,
  Result,
  SettingsFolderName,
} from "@microsoft/teamsfx-api";
import { assert, expect } from "chai";
import fs from "fs-extra";
import "mocha";
import * as os from "os";
import * as path from "path";
import * as sinon from "sinon";
import * as projectSettingsHelper from "../../../src/common/projectSettingsHelper";
import * as tools from "../../../src/common/utils";
import { CallbackRegistry } from "../../../src/core/callback";
import { runWithRetry } from "../../../src/core/middleware/retry";
import { CoreSource, NoProjectOpenedError } from "../../../src/error";
import {
  ConcurrentError,
  FileNotFoundError,
  InvalidProjectError,
  UserCancelError,
} from "../../../src/error/common";
import { MockTools, randomAppName } from "../utils";
import { setTools } from "../../../src/common/globalVars";

describe("runWithRetry", () => {
  const tools = new MockTools();
  setTools(tools);
  it("no retry", async () => {
    let callNum = 0;
    await runWithRetry(
      async () => {
        ++callNum;
        return ok("");
      },
      (result, attempt) => false
    );
    assert.equal(callNum, 1);
  });

  it("no retry with error thrown", async () => {
    let callNum = 0;
    try {
      await runWithRetry(
        async () => {
          ++callNum;
          throw new UserCancelError();
        },
        (result, attempt) => false
      );
      assert.fail("should not reach here");
    } catch (e) {
      assert.isTrue(e instanceof UserCancelError);
    }
    assert.equal(callNum, 1);
  });

  it("retry once", async () => {
    let callNum = 0;
    await runWithRetry(
      async () => {
        ++callNum;
        return ok("");
      },
      (result, attempt) => {
        if (attempt === 1) return true;
        return false;
      }
    );
    assert.equal(callNum, 2);
  });

  it("retry until max", async () => {
    let callNum = 0;
    await runWithRetry(
      async () => {
        ++callNum;
        return ok("");
      },
      (result, attempt) => true
    );
    assert.equal(callNum, 3);
  });
});
