// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { Colors, LogLevel } from "@microsoft/teamsfx-api";
import { expect } from "chai";
import sinon from "sinon";
import { CLILogProvider } from "../../../src/commonlib/log";
import ScreenManager from "../../../src/console/screen";
import { CLILogLevel } from "../../../src/constants";

describe("CLILogProvider", () => {
  const logger = new CLILogProvider();
  logger.setLogLevel(CLILogLevel.debug);
  const sandox = sinon.createSandbox();
  let message = "";

  beforeEach(() => {
    sandox.stub(ScreenManager, "writeLine").callsFake((msg: string) => (message += msg));
  });

  afterEach(() => {
    sandox.restore();
    message = "";
  });

  it("Log - verbose", async () => {
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
    await logger.logInFile(LogLevel.Info, "info");
    expect(message).to.eq("");
  });
  it("OutputSuccess", async () => {
    logger.outputSuccess("success");
    expect(message).to.contain("success");
  });

  it("NecessaryLog - Verbose", async () => {
    logger.necessaryLog(LogLevel.Verbose, "trace");
    expect(message).to.contain("trace");
  });

  it("NecessaryLog - Debug", async () => {
    logger.necessaryLog(LogLevel.Debug, "debug");
    expect(message).to.contain("debug");
  });

  it("NecessaryLog - Info", async () => {
    logger.necessaryLog(LogLevel.Info, "info");
    expect(message).to.contain("info");
  });

  it("NecessaryLog - Info - White", async () => {
    logger.necessaryLog(LogLevel.Info, "info", true);
    expect(message).to.contain("info");
  });

  it("NecessaryLog - Warning", async () => {
    logger.necessaryLog(LogLevel.Warning, "warning");
    expect(message).to.contain("warning");
  });

  it("NecessaryLog - Error", async () => {
    logger.necessaryLog(LogLevel.Error, "error");
    expect(message).to.contain("error");
  });
});
