// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import {
  err,
  FxError,
  IProgressHandler,
  ok,
  Result,
  UserError,
  signedIn,
  signedOut,
} from "@microsoft/teamsfx-api";
import { envUtil, FxCore, HubTypes, VersionCheckRes, VersionState } from "@microsoft/teamsfx-core";
import * as packageJson from "@microsoft/teamsfx-core/build/component/local/packageJsonHelper";
import * as tools from "@microsoft/teamsfx-core/build/common/tools";
import fs from "fs-extra";
import { RestoreFn } from "mocked-env";
import * as path from "path";
import sinon from "sinon";
import * as commonUtils from "../../../../src/cmds/preview/commonUtils";
import * as constants from "../../../../src/cmds/preview/constants";
import * as launch from "../../../../src/cmds/preview/launch";
import PreviewEnv from "../../../../src/cmds/preview/previewEnv";
import { ServiceLogWriter } from "../../../../src/cmds/preview/serviceLogWriter";
import { Task } from "../../../../src/cmds/preview/task";
import cliLogger from "../../../../src/commonlib/log";
import M365TokenInstance from "../../../../src/commonlib/m365Login";
import cliTelemetry from "../../../../src/telemetry/cliTelemetry";
import CLIUIInstance from "../../../../src/userInteraction";
import * as Utils from "../../../../src/utils";
import { expect } from "../../utils";
import * as settingHelper from "@microsoft/teamsfx-core/build/common/projectSettingsHelper";
import { unique } from "underscore";

describe("Preview --env", () => {
  const sandbox = sinon.createSandbox();
  let mockedEnvRestore: RestoreFn = () => {};
  let options: string[] = [];
  let defaultOptions: { [k: string]: any } = {};
  let logs: string[] = [];
  let telemetries: any[] = [];

  beforeEach(() => {
    mockedEnvRestore = () => {};
    options = [];
    defaultOptions = { folder: "./", env: "local" };
    logs = [];
    telemetries = [];
    sandbox.stub(process, "exit");
    sandbox.stub(cliLogger, "necessaryLog").callsFake((lv, msg, white) => {
      logs.push(msg);
    });
    sandbox.stub(cliLogger, "outputInfo").callsFake((message: string) => {
      logs.push(message);
    });
    sandbox.stub(cliLogger, "outputError").callsFake((message: string) => {
      logs.push(message);
    });
    sandbox.stub(cliLogger, "outputSuccess").callsFake((message: string) => {
      logs.push(message);
    });
    sandbox.stub(cliTelemetry, "sendTelemetryEvent").callsFake((eventName, properties) => {
      telemetries.push([eventName, properties]);
    });
    sandbox
      .stub(cliTelemetry, "sendTelemetryErrorEvent")
      .callsFake((eventName, error, properties) => {
        telemetries.push([eventName, error, properties]);
      });
    sandbox.stub(FxCore.prototype, "projectVersionCheck").resolves(
      ok<VersionCheckRes, FxError>({
        isSupport: VersionState.compatible,
        versionSource: "",
        currentVersion: "1.0.0",
        trackingId: "",
      })
    );
  });

  afterEach(() => {
    sandbox.restore();
    mockedEnvRestore();
  });

  it("Preview Command Running - Default", async () => {
    sandbox.stub(settingHelper, "isValidProjectV3").returns(true);
    sandbox.stub(envUtil, "readEnv").resolves(ok({}));
    sandbox.stub(PreviewEnv.prototype, <any>"checkM365Account").resolves(ok({}));
    sandbox.stub(PreviewEnv.prototype, <any>"previewWithManifest").resolves(ok("test-url"));
    sandbox.stub(PreviewEnv.prototype, <any>"detectRunCommand").resolves(ok({}));
    sandbox.stub(PreviewEnv.prototype, <any>"runCommandAsTask").resolves(ok(null));
    sandbox.stub(PreviewEnv.prototype, <any>"launchBrowser").resolves(ok(null));
    const cmd = new PreviewEnv();
    await cmd.runCommand(defaultOptions);
    expect(logs.length >= 1).to.be.true;
    expect(logs[0]).satisfy((l: string) => l.includes("run-command"));
  });

  it("Preview Command Running - outlook", async () => {
    sandbox.stub(settingHelper, "isValidProjectV3").returns(true);
    sandbox.stub(envUtil, "readEnv").resolves(ok({}));
    sandbox.stub(PreviewEnv.prototype, <any>"checkM365Account").resolves(ok({}));
    sandbox.stub(PreviewEnv.prototype, <any>"previewWithManifest").resolves(ok("test-url"));
    sandbox.stub(PreviewEnv.prototype, <any>"detectRunCommand").resolves(ok({}));
    sandbox.stub(PreviewEnv.prototype, <any>"runCommandAsTask").resolves(ok(null));
    sandbox.stub(PreviewEnv.prototype, <any>"launchBrowser").resolves(ok(null));
    const cmd = new PreviewEnv();
    await cmd.runCommand({
      ...defaultOptions,
      ["m365-host"]: "outlook",
      ["browser-arg"]: ["--guest"],
      ["open-only"]: true,
    });
    expect(logs.length >= 0).to.be.true;
  });

  it("Preview Command Running - office", async () => {
    sandbox.stub(settingHelper, "isValidProjectV3").returns(true);
    sandbox.stub(envUtil, "readEnv").resolves(ok({}));
    sandbox.stub(PreviewEnv.prototype, <any>"checkM365Account").resolves(ok({}));
    sandbox.stub(PreviewEnv.prototype, <any>"previewWithManifest").resolves(ok("test-url"));
    sandbox.stub(PreviewEnv.prototype, <any>"detectRunCommand").resolves(ok({}));
    sandbox.stub(PreviewEnv.prototype, <any>"runCommandAsTask").resolves(ok(null));
    sandbox.stub(PreviewEnv.prototype, <any>"launchBrowser").resolves(ok(null));

    const cmd = new PreviewEnv();
    await cmd.runCommand({
      ...defaultOptions,
      env: "dev",
      ["m365-host"]: "office",
      ["browser-arg"]: ["--guest"],
    });

    expect(logs.length >= 0).to.be.true;
  });

  it("Preview Command Running - workspace not supported error", async () => {
    sandbox.stub(settingHelper, "isValidProjectV3").returns(false);

    const cmd = new PreviewEnv();
    const result = await cmd.runCommand(defaultOptions);

    expect(result.isErr()).to.be.true;
    expect((result as any).error.name).equals("WorkspaceNotSupported");
  });

  it("Preview Command Running - load envs error", async () => {
    sandbox.stub(settingHelper, "isValidProjectV3").returns(true);
    sandbox.stub(envUtil, "readEnv").resolves(err({ foo: "bar" } as any));

    const cmd = new PreviewEnv();
    const result = await cmd.runCommand(defaultOptions);

    expect(result.isErr()).to.be.true;
    expect((result as any).error).to.deep.equal({ foo: "bar" });
  });

  it("Preview Command Running - check account error", async () => {
    sandbox.stub(settingHelper, "isValidProjectV3").returns(true);
    sandbox.stub(envUtil, "readEnv").resolves(ok({}));
    sandbox
      .stub(PreviewEnv.prototype, <any>"checkM365Account")
      .resolves(err({ foo: "bar" } as any));

    const cmd = new PreviewEnv();
    const result = await cmd.runCommand(defaultOptions);

    expect(result.isErr()).to.be.true;
    expect((result as any).error).to.deep.equal({ foo: "bar" });
  });

  it("Preview Command Running - previewWithManifest error", async () => {
    sandbox.stub(settingHelper, "isValidProjectV3").returns(true);
    sandbox.stub(envUtil, "readEnv").resolves(ok({}));
    sandbox.stub(PreviewEnv.prototype, <any>"checkM365Account").resolves(ok({}));
    sandbox
      .stub(PreviewEnv.prototype, <any>"previewWithManifest")
      .resolves(err({ foo: "bar" } as any));

    const cmd = new PreviewEnv();
    const result = await cmd.runCommand(defaultOptions);

    expect(result.isErr()).to.be.true;
    expect((result as any).error).to.deep.equal({ foo: "bar" });
  });

  it("Preview Command Running - detect run command error", async () => {
    sandbox.stub(settingHelper, "isValidProjectV3").returns(true);
    sandbox.stub(envUtil, "readEnv").resolves(ok({}));
    sandbox.stub(PreviewEnv.prototype, <any>"checkM365Account").resolves(ok({}));
    sandbox.stub(PreviewEnv.prototype, <any>"previewWithManifest").resolves(ok("test-url"));
    sandbox
      .stub(PreviewEnv.prototype, <any>"detectRunCommand")
      .resolves(err({ foo: "bar" } as any));

    const cmd = new PreviewEnv();
    const result = await cmd.runCommand(defaultOptions);

    expect(result.isErr()).to.be.true;
    expect((result as any).error).to.deep.equal({ foo: "bar" });
  });

  it("Preview Command Running - run task error", async () => {
    sandbox.stub(settingHelper, "isValidProjectV3").returns(true);
    sandbox.stub(envUtil, "readEnv").resolves(ok({}));
    sandbox.stub(PreviewEnv.prototype, <any>"checkM365Account").resolves(ok({}));
    sandbox.stub(PreviewEnv.prototype, <any>"previewWithManifest").resolves(ok("test-url"));
    sandbox
      .stub(PreviewEnv.prototype, <any>"detectRunCommand")
      .resolves(ok({ runCommand: "npm start" }));
    sandbox
      .stub(PreviewEnv.prototype, <any>"runCommandAsTask")
      .resolves(err({ foo: "bar" } as any));

    const cmd = new PreviewEnv();
    const result = await cmd.runCommand(defaultOptions);
    expect(result.isErr()).to.be.true;
    expect((result as any).error).to.deep.equal({ foo: "bar" });
  });

  it("Preview Command Running - launch browser error", async () => {
    sandbox.stub(settingHelper, "isValidProjectV3").returns(true);
    sandbox.stub(envUtil, "readEnv").resolves(ok({}));
    sandbox.stub(PreviewEnv.prototype, <any>"checkM365Account").resolves(ok({}));
    sandbox.stub(PreviewEnv.prototype, <any>"previewWithManifest").resolves(ok("test-url"));
    sandbox.stub(PreviewEnv.prototype, <any>"detectRunCommand").resolves(ok({}));
    sandbox.stub(PreviewEnv.prototype, <any>"runCommandAsTask").resolves(ok(null));
    sandbox.stub(PreviewEnv.prototype, <any>"launchBrowser").resolves(err({ foo: "bar" } as any));
    const cmd = new PreviewEnv();
    const result = await cmd.runCommand(defaultOptions);
    expect(result.isErr()).to.be.true;
    expect((result as any).error).to.deep.equal({ foo: "bar" });
  });
});

describe("PreviewEnv Steps", () => {
  const sandbox = sinon.createSandbox();
  let mockedEnvRestore: RestoreFn = () => {};
  let logs: string[] = [];
  let telemetries: any[] = [];

  // wrapper class to expose protected functions
  class PreviewEnvTest extends PreviewEnv {
    public checkM365Account(appTenantId?: string): Promise<
      Result<
        {
          tenantId?: string;
          loginHint?: string;
        },
        FxError
      >
    > {
      return super.checkM365Account(appTenantId);
    }

    public async detectRunCommand(projectPath: string): Promise<
      Result<
        {
          runCommand: string;
        },
        FxError
      >
    > {
      return await super.detectRunCommand(projectPath);
    }

    public previewWithManifest(
      projectPath: string,
      env: string,
      hub: HubTypes,
      manifestFilePath: string
    ): Promise<Result<string, FxError>> {
      return super.previewWithManifest(projectPath, env, hub, manifestFilePath);
    }

    public runCommandAsTask(
      projectPath: string,
      runCommand: string,
      runningPatternRegex: RegExp,
      execPath: string
    ): Promise<Result<null, FxError>> {
      return super.runCommandAsTask(projectPath, runCommand, runningPatternRegex, execPath);
    }

    public launchBrowser(
      env: string,
      hub: HubTypes,
      url: string,
      browser: constants.Browser,
      browserArgs: string[]
    ): Promise<Result<null, FxError>> {
      return super.launchBrowser(env, hub, url, browser, browserArgs);
    }

    public launchDesktopClient(
      env: string,
      url: string,
      browser: constants.Browser,
      browserArgs: string[]
    ): Promise<Result<null, FxError>> {
      return super.launchDesktopClient(env, url, browser, browserArgs);
    }

    public getRunningTasks() {
      return this.runningTasks;
    }
  }

  beforeEach(() => {
    mockedEnvRestore = () => {};
    logs = [];
    telemetries = [];
    sandbox.stub(cliLogger, "necessaryLog").callsFake((lv, msg, white) => {
      logs.push(msg);
    });
    sandbox.stub(cliTelemetry, "sendTelemetryEvent").callsFake((eventName, properties) => {
      telemetries.push([eventName, properties]);
    });
    sandbox
      .stub(cliTelemetry, "sendTelemetryErrorEvent")
      .callsFake((eventName, error, properties) => {
        telemetries.push([eventName, error, properties]);
      });
    sandbox.stub(CLIUIInstance, "createProgressBar").returns(new MockProgressHandler());
  });

  afterEach(() => {
    sandbox.restore();
    mockedEnvRestore();
  });

  it("checkM365Account - signin", async () => {
    const token = "test-token";
    const tenantId = "test-tenant-id";
    const upn = "test-user";
    sandbox.stub(M365TokenInstance, "getStatus").returns(
      Promise.resolve(
        ok({
          status: signedIn,
          token: token,
          accountInfo: {
            tid: tenantId,
            upn: upn,
          },
        })
      )
    );
    sandbox.stub(tools, "getSideloadingStatus").resolves(true);

    const previewEnv = new PreviewEnvTest();
    const accountRes = await previewEnv.checkM365Account(undefined);
    expect(accountRes.isOk()).to.be.true;
    const account = (accountRes as any).value;
    expect(account.tenantId).equals(tenantId);
    expect(account.loginHint).equals(upn);
  });

  it("checkM365Account - signout", async () => {
    const token = "test-token";
    const tenantId = "test-tenant-id";
    const upn = "test-user";
    const getStatusStub = sandbox.stub(M365TokenInstance, "getStatus");
    getStatusStub.onCall(0).resolves(
      ok({
        status: signedOut,
      })
    );
    getStatusStub.onCall(1).resolves(
      ok({
        status: signedIn,
        token: token,
        accountInfo: {
          tid: tenantId,
          upn: upn,
        },
      })
    );
    sandbox.stub(M365TokenInstance, "getAccessToken").resolves(ok(token));
    sandbox.stub(tools, "getSideloadingStatus").resolves(true);

    const previewEnv = new PreviewEnvTest();
    const accountRes = await previewEnv.checkM365Account(undefined);
    expect(accountRes.isOk()).to.be.true;
    const account = (accountRes as any).value;
    expect(account.tenantId).equals(tenantId);
    expect(account.loginHint).equals(upn);
  });

  it("checkM365Account - no sideloading permission", async () => {
    const token = "test-token";
    const tenantId = "test-tenant-id";
    const upn = "test-user";
    sandbox.stub(M365TokenInstance, "getStatus").returns(
      Promise.resolve(
        ok({
          status: signedIn,
          token: token,
          accountInfo: {
            tid: tenantId,
            upn: upn,
          },
        })
      )
    );
    sandbox.stub(tools, "getSideloadingStatus").resolves(false);

    const previewEnv = new PreviewEnvTest();
    const accountRes = await previewEnv.checkM365Account(undefined);
    expect(accountRes.isErr()).to.be.true;
    const error = (accountRes as any).error;
    // eslint-disable-next-line no-secrets/no-secrets
    expect(error.name).equals("PrerequisitesValidationM365AccountError");
    expect(error.message).satisfy((m: string) => m.includes("custom app upload permission"));
  });

  it("detectRunCommand - node", async () => {
    sandbox.stub(fs, "pathExists").resolves(true);
    sandbox.stub(fs, "readdir").resolves([]);
    // eslint-disable-next-line no-secrets/no-secrets
    sandbox.stub(packageJson, "loadTeamsFxDevScript").resolves("test");

    const previewEnv = new PreviewEnvTest();
    const runCommandRes = await previewEnv.detectRunCommand("./");
    expect(runCommandRes.isOk()).to.be.true;
    const runCommand = (runCommandRes as any).value;
    expect(runCommand.runCommand).equals("npm run dev:teamsfx");
  });

  it("detectRunCommand - .net web", async () => {
    sandbox.stub(fs, "pathExists").resolves(false);
    sandbox.stub(fs, "readdir").resolves(["test.csproj"]);
    sandbox.stub(fs, "readFile").resolves(
      Buffer.from(
        `
  <Project Sdk="Microsoft.NET.Sdk.Web">
    <PropertyGroup>
      <TargetFramework>net6.0</TargetFramework>
    </PropertyGroup>
    <ItemGroup>
      <ProjectCapability Include="TeamsFx" />
    </ItemGroup>
  </Project>
  `
      )
    );

    const previewEnv = new PreviewEnvTest();
    const runCommandRes = await previewEnv.detectRunCommand("./");
    expect(runCommandRes.isOk()).to.be.true;
    const runCommand = (runCommandRes as any).value;
    expect(runCommand.runCommand).equals("dotnet run");
  });

  it("detectRunCommand - .net func", async () => {
    sandbox.stub(fs, "pathExists").resolves(false);
    sandbox.stub(fs, "readdir").resolves(["test.csproj"]);
    sandbox.stub(fs, "readFile").resolves(
      Buffer.from(
        // eslint-disable-next-line no-secrets/no-secrets
        `
  <Project Sdk="Microsoft.NET.Sdk">
    <PropertyGroup>
      <TargetFramework>net6.0</TargetFramework>
    </PropertyGroup>
    <ItemGroup>
      <ProjectCapability Include="TeamsFx" />
    </ItemGroup>
    <ItemGroup>
      <PackageReference Include="Microsoft.NET.Sdk.Functions" Version="4.1.3" />
    </ItemGroup>
  </Project>
  `
      )
    );

    const previewEnv = new PreviewEnvTest();
    const runCommandRes = await previewEnv.detectRunCommand("./");
    expect(runCommandRes.isOk()).to.be.true;
    const runCommand = (runCommandRes as any).value;
    expect(runCommand.runCommand).equals("func start");
  });

  it("previewWithManifest - previewWithManifest error", async () => {
    sandbox.stub(FxCore.prototype, "previewWithManifest").resolves(err({ foo: "bar" } as any));

    const previewEnv = new PreviewEnvTest();
    const result = await previewEnv.previewWithManifest(
      "./",
      "local",
      HubTypes.teams,
      "test-manifest-path"
    );
    expect(result.isErr()).to.be.true;
    expect((result as any).error).to.deep.equal({ foo: "bar" });
  });

  it("previewWithManifest - ok", async () => {
    sandbox.stub(FxCore.prototype, "previewWithManifest").resolves(ok("test-url"));

    const previewEnv = new PreviewEnvTest();
    const result = await previewEnv.previewWithManifest(
      "./",
      "local",
      HubTypes.teams,
      "test-manifest-path"
    );
    expect(result.isOk()).to.be.true;
  });

  it("runCommandAsTask - ok", async () => {
    sandbox
      .stub(commonUtils, "createTaskStartCb")
      .returns((a0: any, a1: any) => new Promise((res, rej) => res()));
    sandbox
      .stub(commonUtils, "createTaskStopCb")
      .returns((a0: any, a1: any, a2: any, a3: any) => new Promise((res, rej) => res(null)));
    sandbox.stub(ServiceLogWriter.prototype, "init").resolves();
    sandbox.stub(Task.prototype, "waitFor").resolves(ok({ foo: "bar" } as any));

    const previewEnv = new PreviewEnvTest();
    const taskRes = await previewEnv.runCommandAsTask(
      "./",
      "npm start",
      /done/i,
      "./devTools/func"
    );
    expect(taskRes.isOk()).to.be.true;
    const tasks = previewEnv.getRunningTasks();
    expect(tasks.length).equals(1);
    expect((tasks[0] as any).taskTitle).equals("Run Command");
    expect((tasks[0] as any).command).equals("npm start");
  });

  it("runCommandAsTask - customize exec path", async () => {
    sandbox
      .stub(commonUtils, "createTaskStartCb")
      .returns((a0: any, a1: any) => new Promise((res, rej) => res()));
    sandbox
      .stub(commonUtils, "createTaskStopCb")
      .returns((a0: any, a1: any, a2: any, a3: any) => new Promise((res, rej) => res(null)));
    sandbox.stub(ServiceLogWriter.prototype, "init").resolves();
    sandbox.stub(Task.prototype, "waitFor").resolves(ok({ foo: "bar" } as any));

    const previewEnv = new PreviewEnvTest();
    const taskRes = await previewEnv.runCommandAsTask(
      "./",
      "npm start",
      /done/i,
      `./devTools/func${path.delimiter}${path.resolve(`./devTools/func1`)}`
    );
    expect(taskRes.isOk()).to.be.true;
    const tasks = previewEnv.getRunningTasks();
    expect(tasks.length).equals(1);
    expect((tasks[0] as any).taskTitle).equals("Run Command");
    expect((tasks[0] as any).command).equals("npm start");
    expect((tasks[0] as any).options.env.PATH).include(
      `${path.resolve("./devTools/func")}${path.delimiter}${path.resolve("./devTools/func1")}${
        path.delimiter
      }`
    );
  });

  it("launchBrowser - teams", async () => {
    sandbox.stub(launch, "openHubWebClientNew").resolves();

    const previewEnv = new PreviewEnvTest();
    const openRes = await previewEnv.launchBrowser(
      "local",
      HubTypes.teams,
      "test-url",
      constants.Browser.default,
      []
    );
    expect(openRes.isOk()).to.be.true;
    expect(logs.length).equals(1);
  });

  it("launchBrowser: outlook", async () => {
    sandbox.stub(launch, "openHubWebClientNew").resolves();

    const previewEnv = new PreviewEnvTest();
    const openRes = await previewEnv.launchBrowser(
      "local",
      HubTypes.outlook,
      "test-url",
      constants.Browser.default,
      []
    );
    expect(openRes.isOk()).to.be.true;
    expect(logs.length).equals(1);
  });

  it("launchDesktopClient - without accountInfo", async () => {
    sandbox.stub(launch, "openTeamsDesktopClient").resolves();
    sandbox.stub(M365TokenInstance, "getStatus").returns(
      Promise.resolve(
        ok({
          status: signedIn,
          token: "token",
        })
      )
    );

    const previewEnv = new PreviewEnvTest();
    const openRes = await previewEnv.launchDesktopClient(
      "local",
      "test-url",
      constants.Browser.default,
      []
    );
    expect(openRes.isOk()).to.be.true;
  });

  it("launchDesktopClient - without unique_name", async () => {
    sandbox.stub(launch, "openTeamsDesktopClient").resolves();
    sandbox.stub(M365TokenInstance, "getStatus").returns(
      Promise.resolve(
        ok({
          status: signedIn,
          token: "token",
          accountInfo: {
            tid: "tid",
            upn: "upn",
          },
        })
      )
    );

    const previewEnv = new PreviewEnvTest();
    const openRes = await previewEnv.launchDesktopClient(
      "local",
      "test-url",
      constants.Browser.default,
      []
    );
    expect(openRes.isOk()).to.be.true;
  });

  it("launchDesktopClient - happy path", async () => {
    sandbox.stub(launch, "openTeamsDesktopClient").resolves();
    sandbox.stub(M365TokenInstance, "getStatus").returns(
      Promise.resolve(
        ok({
          status: signedIn,
          token: "token",
          accountInfo: {
            tid: "tid",
            upn: "upn",
            unique_name: "unique_name",
          },
        })
      )
    );

    const previewEnv = new PreviewEnvTest();
    const openRes = await previewEnv.launchDesktopClient(
      "local",
      "test-url",
      constants.Browser.default,
      []
    );
    expect(openRes.isOk()).to.be.true;
  });

  it("launchDesktopClient - without user information", async () => {
    sandbox.stub(launch, "openTeamsDesktopClient").resolves();
    sandbox.stub(M365TokenInstance, "getStatus").resolves(err(new UserError("", "", "", "")));

    const previewEnv = new PreviewEnvTest();
    const openRes = await previewEnv.launchDesktopClient(
      "local",
      "test-url",
      constants.Browser.default,
      []
    );
    expect(openRes.isOk()).to.be.true;
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
