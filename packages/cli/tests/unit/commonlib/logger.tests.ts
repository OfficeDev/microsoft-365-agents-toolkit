// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { Colors, LogLevel } from "@microsoft/teamsfx-api";
import { expect } from "chai";
import sinon from "sinon";
import ScreenManager from "../../../src/console/screen";
import { logger } from "../../../src/commonlib/logger";

describe("CLILogger", () => {
  logger.logLevel = LogLevel.Debug;
  const sandox = sinon.createSandbox();
  let message = "";

  beforeEach(() => {
    sandox.stub(ScreenManager, "writeLine").callsFake((msg: string) => (message += msg));
  });

  afterEach(() => {
    sandox.restore();
    message = "";
  });

  it("Log - Verbose", async () => {
    logger.verbose("verbose");
    expect(message).to.contain("verbose");
  });

  it("Log - Debug", async () => {
    logger.debug("debug");
    expect(message).to.contain("debug");
  });

  it("Log - Info", async () => {
    logger.info("info");
    expect(message).to.contain("info");
  });
  it("Log - Info", async () => {
    logger.info([{ content: "infocolor", color: Colors.WHITE }]);
    expect(message).to.contain("infocolor");
  });
  it("Log - Warning", async () => {
    logger.warning("warning");
    expect(message).to.contain("warning");
  });

  it("Log - Error", async () => {
    logger.error("error");
    expect(message).to.contain("error");
  });

  it("logInFile", async () => {
    logger.logInFile(LogLevel.Info, "info");
    expect(message).to.eq("");
  });
  it("OutputSuccess", async () => {
    logger.outputSuccess("success");
    expect(message).to.contain("success");
  });
});
