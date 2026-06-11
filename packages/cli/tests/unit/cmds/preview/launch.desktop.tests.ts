// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { IProgressHandler } from "@microsoft/teamsfx-api";
import cp from "child_process";
import * as sinon from "sinon";
import * as commonUtils from "../../../../src/cmds/preview/commonUtils";
import { Browser } from "../../../../src/cmds/preview/constants";
import { openTeamsDesktopClient } from "../../../../src/cmds/preview/launch";
import cliLogger from "../../../../src/commonlib/log";
import cliTelemetry from "../../../../src/telemetry/cliTelemetry";
import CLIUIInstance from "../../../../src/userInteraction";
import { expect } from "../../utils";

describe("launch openTeamsDesktopClient", () => {
  const sandbox = sinon.createSandbox();
  afterEach(() => {
    sandbox.restore();
  });

  let telemetries: any[] = [];
  const telemetryProperties = {
    key1: "value1",
    key2: "value2",
  };

  beforeEach(() => {
    telemetries = [];

    sandbox.stub(process.stdout, "write").returns(true as any);
    sandbox.stub(process.stderr, "write").returns(true as any);

    sandbox.stub(cliTelemetry, "sendTelemetryEvent").callsFake((eventName, properties) => {
      telemetries.push([eventName, properties]);
    });
    sandbox
      .stub(cliTelemetry, "sendTelemetryErrorEvent")
      .callsFake((eventName, error, properties) => {
        telemetries.push([eventName, error, properties]);
      });
    sandbox.stub(cliLogger, "necessaryLog").callsFake(() => {});
    sandbox.stub(CLIUIInstance, "createProgressBar").returns(new MockProgressHandler());
    sandbox.stub(cp, "exec");
  });

  it("happy path windows", async () => {
    sandbox.stub(process, "platform").value("win32");
    await openTeamsDesktopClient("http://test-url", "username", Browser.default);
    expect(telemetries.length).to.deep.equals(0);
  });

  it("happy path mac", async () => {
    sandbox.stub(process, "platform").value("darwin");
    await openTeamsDesktopClient("http://test-url", "username", Browser.default);
    expect(telemetries.length).to.deep.equals(0);
  });

  it("happy path windows - with telemetry", async () => {
    sandbox.stub(process, "platform").value("win32");
    await openTeamsDesktopClient(
      "http://test-url",
      "username",
      Browser.default,
      [],
      telemetryProperties
    );
    expect(telemetries.length).to.deep.equals(2);
  });

  it("happy path others", async () => {
    sandbox.stub(process, "platform").value("linux");
    sandbox
      .stub(commonUtils, "openBrowser")
      .callsFake(async (browser, url, browserArguments) => {});
    await openTeamsDesktopClient("http://test-url", "username", Browser.default, ["test"]);
    expect(telemetries.length).to.deep.equals(0);
  });

  it("openBrowser error", async () => {
    sandbox.stub(process, "platform").value("linux");
    sandbox.stub(commonUtils, "openBrowser").throws();
    await openTeamsDesktopClient("http://test-url", "username", Browser.default);
    expect(telemetries.length).to.deep.equals(0);
  });
});

class MockProgressHandler implements IProgressHandler {
  start(detail?: string): Promise<void> {
    return Promise.resolve();
  }
  next(detail?: string): Promise<void> {
    return Promise.resolve();
  }
  end(success: boolean): Promise<void> {
    return Promise.resolve();
  }
}
