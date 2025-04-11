import {
  err,
  FxError,
  LogProvider,
  ok,
  Result,
  SystemError,
  TeamsAppManifest,
} from "@microsoft/teamsfx-api";
import { assert } from "chai";
import "mocha";
import sinon from "sinon";
import { TelemetryEvent, TelemetryProperty } from "../../../src/common/telemetry";
import {
  DriverInstance,
  ExecutionResult,
  ProjectModel,
} from "../../../src/component/configManager/interface";
import { yamlParser } from "../../../src/component/configManager/parser";
import { DriverContext } from "../../../src/component/driver/interface/commonArgs";
import { metadataUtil } from "../../../src/component/utils/metadataUtil";
import { setTools } from "../../../src/common/globalVars";
import { MockTools } from "../../core/utils";
import { createHash, Hash } from "crypto";
import { ExecutionResult as DriverResult } from "../../../src/component/driver/interface/stepDriver";
import { metadataGraphPermissionUtil } from "../../../src/component/utils/metadataGraphPermssion";

function mockedResolveDriverInstances(log: LogProvider): Result<DriverInstance[], FxError> {
  return ok([
    {
      uses: "arm/deploy",
      with: undefined,
      instance: {
        execute: async (args: unknown, context: DriverContext): Promise<DriverResult> => {
          return { result: ok(new Map<string, string>()), summaries: [] };
        },
      },
    },
  ]);
}

describe("metadata util", () => {
  const sandbox = sinon.createSandbox();
  const mockedError = new SystemError("mockedSource", "mockedError", "mockedMessage");
  const mockProjectModel: ProjectModel = {
    version: "1.0.0",
    additionalMetadata: {
      sampleTag: "testRepo:testSample",
    },
    registerApp: {
      name: "registerApp",
      driverDefs: [
        {
          uses: "arm/deploy",
          with: undefined,
        },
        {
          uses: "teamsApp/create",
          with: undefined,
        },
      ],
      resolvePlaceholders: () => {
        return ["AZURE_SUBSCRIPTION_ID", "AZURE_RESOURCE_GROUP_NAME"];
      },
      execute: async (ctx: DriverContext): Promise<ExecutionResult> => {
        return { result: ok(new Map()), summaries: [] };
      },
      resolveDriverInstances: mockedResolveDriverInstances,
    },
    environmentFolderPath: "./envs",
  };
  let tools: MockTools;

  beforeEach(() => {
    tools = new MockTools();
    setTools(tools);
    sandbox.stub(metadataGraphPermissionUtil, "parseAadManifest").resolves();
  });

  afterEach(() => {
    sandbox.restore();
  });

  it("should return YamlParsingError", async () => {
    sandbox.stub(yamlParser, "parse").resolves(err(mockedError));
    const result = await metadataUtil.parse(".", "dev");
    assert(result.isErr() && result.error.name === "mockedError");
  });

  it("local config file", async () => {
    sandbox.stub(yamlParser, "parse").resolves(ok(mockProjectModel));
    const spy = sandbox.spy(tools.telemetryReporter, "sendTelemetryEvent");
    const result = await metadataUtil.parse(".", "local");
    assert.isTrue(
      spy.calledOnceWith(TelemetryEvent.MetaData, {
        [TelemetryProperty.YmlSchemaVersion]: "1.0.0",
        "configureApp.actions": "",
        "deploy.actions": "",
        "share.actions": "",
        "provision.actions": "",
        "publish.actions": "",
        "registerApp.actions": "armdeploy,teamsAppcreate",
        [TelemetryProperty.YmlName]: "teamsapplocalyml",
        [TelemetryProperty.SampleAppName]: "testRepo:testSample",
      })
    );
    assert(result.isOk());
  });

  it("dev config file", async () => {
    sandbox.stub(yamlParser, "parse").resolves(ok(mockProjectModel));
    const spy = sandbox.spy(tools.telemetryReporter, "sendTelemetryEvent");
    const result = await metadataUtil.parse(".", "dev");
    assert.isTrue(
      spy.calledOnceWith(TelemetryEvent.MetaData, {
        [TelemetryProperty.YmlSchemaVersion]: "1.0.0",
        "configureApp.actions": "",
        "deploy.actions": "",
        "share.actions": "",
        "provision.actions": "",
        "publish.actions": "",
        "registerApp.actions": "armdeploy,teamsAppcreate",
        [TelemetryProperty.YmlName]: "teamsappyml",
        [TelemetryProperty.SampleAppName]: "testRepo:testSample",
      })
    );
    assert(result.isOk());
  });

  it("should normalize @/\\. in sampleTag", async () => {
    sandbox.stub(yamlParser, "parse").resolves(
      ok({
        ...mockProjectModel,
        additionalMetadata: { sampleTag: "Hello@world/this\\is.a.sample" },
      })
    );
    const spy = sandbox.spy(tools.telemetryReporter, "sendTelemetryEvent");
    const result = await metadataUtil.parse(".", "dev");
    assert.isTrue(
      spy.calledOnceWith(TelemetryEvent.MetaData, {
        [TelemetryProperty.YmlSchemaVersion]: "1.0.0",
        "configureApp.actions": "",
        "deploy.actions": "",
        "share.actions": "",
        "provision.actions": "",
        "publish.actions": "",
        "registerApp.actions": "armdeploy,teamsAppcreate",
        [TelemetryProperty.YmlName]: "teamsappyml",
        [TelemetryProperty.SampleAppName]: "Hello_world_this_is_a_sample",
      })
    );
    assert(result.isOk());
  });

  it("should send empty sample-app-name if additionalMetadata is undefined", async () => {
    sandbox.stub(yamlParser, "parse").resolves(
      ok({
        ...mockProjectModel,
        additionalMetadata: undefined,
      })
    );
    const spy = sandbox.spy(tools.telemetryReporter, "sendTelemetryEvent");
    const result = await metadataUtil.parse(".", "dev");
    assert.isTrue(
      spy.calledOnceWith(TelemetryEvent.MetaData, {
        [TelemetryProperty.YmlSchemaVersion]: "1.0.0",
        "configureApp.actions": "",
        "deploy.actions": "",
        "share.actions": "",
        "provision.actions": "",
        "publish.actions": "",
        "registerApp.actions": "armdeploy,teamsAppcreate",
        [TelemetryProperty.YmlName]: "teamsappyml",
        [TelemetryProperty.SampleAppName]: "",
      })
    );
    assert(result.isOk());
  });

  it("no sample tag", async () => {
    sandbox.stub(yamlParser, "parse").resolves(ok({ ...mockProjectModel, additionalMetadata: {} }));
    const spy = sandbox.spy(tools.telemetryReporter, "sendTelemetryEvent");
    const result = await metadataUtil.parse(".", "dev");
    assert.isTrue(
      spy.calledOnceWith(TelemetryEvent.MetaData, {
        [TelemetryProperty.YmlSchemaVersion]: "1.0.0",
        "configureApp.actions": "",
        "deploy.actions": "",
        "share.actions": "",
        "provision.actions": "",
        "publish.actions": "",
        "registerApp.actions": "armdeploy,teamsAppcreate",
        [TelemetryProperty.YmlName]: "teamsappyml",
        [TelemetryProperty.SampleAppName]: "",
      })
    );
    assert(result.isOk());
  });

  it("parseManifest with empty manifest", () => {
    const spy = sandbox.spy(tools.telemetryReporter, "sendTelemetryEvent");
    metadataUtil.parseManifest({} as any);

    assert.isTrue(
      spy.calledOnceWith(TelemetryEvent.MetaData, {
        "manifest.id": "",
        "manifest.version": "",
        "manifest.manifestVersion": "",
        "manifest.bots": "",
        "manifest.composeExtensions": "",
        "manifest.staticTabs.contentUrl": "",
        "manifest.configurableTabs.configurationUrl": "",
        "manifest.webApplicationInfo.id": "",
        "manifest.extensions": "false",
      })
    );
  });

  it("parseManifest with full manifest", () => {
    const manifest: any = {
      id: "test-id",
      version: "1.0",
      manifestVersion: "1.0",
      bots: [{ botId: "bot1" }, { botId: "bot2" }],
      composeExtensions: [{ botId: "bot1" }, { botId: "bot2" }],
      staticTabs: [
        { contentUrl: "https://example.com/tab1" },
        { contentUrl: "https://example.com/tab2" },
      ],
      configurableTabs: [{ configurationUrl: "https://example.com/config1" }],
      webApplicationInfo: { id: "web-app-id" },
      extensions: [{}],
    };
    const spy = sandbox.spy(tools.telemetryReporter, "sendTelemetryEvent");
    // eslint-disable-next-line @typescript-eslint/ban-ts-comment
    // @ts-ignore
    const hashSpy = sandbox.spy(Hash.prototype);

    manifest.extensions = [{}];
    metadataUtil.parseManifest(manifest as any);
    // eslint-disable-next-line @typescript-eslint/ban-ts-comment
    // @ts-ignore
    assert.isTrue(hashSpy.update.called);
    // eslint-disable-next-line @typescript-eslint/ban-ts-comment
    // @ts-ignore
    assert.isTrue(hashSpy.digest.calledWith("hex"));
    assert.isTrue(
      spy.calledOnceWith(TelemetryEvent.MetaData, {
        "manifest.id": "test-id",
        "manifest.version": "1.0",
        "manifest.manifestVersion": "1.0",
        "manifest.bots": "bot1,bot2",
        "manifest.composeExtensions": "bot1,bot2",
        "manifest.staticTabs.contentUrl": `${[
          createHash("sha256").update(manifest.staticTabs[0].contentUrl).digest("hex"),
          createHash("sha256").update(manifest.staticTabs[1].contentUrl).digest("hex"),
        ].toString()}`,
        "manifest.configurableTabs.configurationUrl": `${createHash("sha256")
          .update(manifest.configurableTabs[0].configurationUrl)
          .digest("hex")}`,
        "manifest.webApplicationInfo.id": "web-app-id",
        "manifest.extensions": "true",
      })
    );

    // If extensions is empty, it should report false in telemetry event
    manifest.extensions = [];
    metadataUtil.parseManifest(manifest as any);
    // eslint-disable-next-line @typescript-eslint/ban-ts-comment
    // @ts-ignore
    assert.isTrue(hashSpy.update.called);
    // eslint-disable-next-line @typescript-eslint/ban-ts-comment
    // @ts-ignore
    assert.isTrue(hashSpy.digest.calledWith("hex"));
    assert.isTrue(
      spy.calledWith(TelemetryEvent.MetaData, {
        "manifest.id": "test-id",
        "manifest.version": "1.0",
        "manifest.manifestVersion": "1.0",
        "manifest.bots": "bot1,bot2",
        "manifest.composeExtensions": "bot1,bot2",
        "manifest.staticTabs.contentUrl": `${[
          createHash("sha256").update(manifest.staticTabs[0].contentUrl).digest("hex"),
          createHash("sha256").update(manifest.staticTabs[1].contentUrl).digest("hex"),
        ].toString()}`,
        "manifest.configurableTabs.configurationUrl": `${createHash("sha256")
          .update(manifest.configurableTabs[0].configurationUrl)
          .digest("hex")}`,
        "manifest.webApplicationInfo.id": "web-app-id",
        "manifest.extensions": "false",
      })
    );

    // If extensions is undefined, it should report false in telemetry event
    manifest.extensions = undefined;
    metadataUtil.parseManifest(manifest as any);
    // eslint-disable-next-line @typescript-eslint/ban-ts-comment
    // @ts-ignore
    assert.isTrue(hashSpy.update.called);
    // eslint-disable-next-line @typescript-eslint/ban-ts-comment
    // @ts-ignore
    assert.isTrue(hashSpy.digest.calledWith("hex"));
    assert.isTrue(
      spy.calledWith(TelemetryEvent.MetaData, {
        "manifest.id": "test-id",
        "manifest.version": "1.0",
        "manifest.manifestVersion": "1.0",
        "manifest.bots": "bot1,bot2",
        "manifest.composeExtensions": "bot1,bot2",
        "manifest.staticTabs.contentUrl": `${[
          createHash("sha256").update(manifest.staticTabs[0].contentUrl).digest("hex"),
          createHash("sha256").update(manifest.staticTabs[1].contentUrl).digest("hex"),
        ].toString()}`,
        "manifest.configurableTabs.configurationUrl": `${createHash("sha256")
          .update(manifest.configurableTabs[0].configurationUrl)
          .digest("hex")}`,
        "manifest.webApplicationInfo.id": "web-app-id",
        "manifest.extensions": "false",
      })
    );
  });

  it("parseManifest with undefined urls", () => {
    const manifest: any = {
      id: "test-id",
      version: "1.0",
      manifestVersion: "1.0",
      bots: [{ botId: "bot1" }, { botId: "bot2" }],
      composeExtensions: [{ botId: "bot1" }, { botId: "bot2" }],
      staticTabs: [{ contentUrl: undefined }, { contentUrl: undefined }],
      configurableTabs: [{ configurationUrl: undefined }],
      webApplicationInfo: { id: "web-app-id" },
    };
    const spy = sandbox.spy(tools.telemetryReporter, "sendTelemetryEvent");
    // eslint-disable-next-line @typescript-eslint/ban-ts-comment
    // @ts-ignore
    const hashSpy = sandbox.spy(Hash.prototype);

    metadataUtil.parseManifest(manifest as any);
    // eslint-disable-next-line @typescript-eslint/ban-ts-comment
    // @ts-ignore
    assert.isTrue(hashSpy.update.notCalled);
    // eslint-disable-next-line @typescript-eslint/ban-ts-comment
    // @ts-ignore
    assert.isTrue(hashSpy.digest.notCalled);
    assert.isTrue(
      spy.calledOnceWith(TelemetryEvent.MetaData, {
        "manifest.id": "test-id",
        "manifest.version": "1.0",
        "manifest.manifestVersion": "1.0",
        "manifest.bots": "bot1,bot2",
        "manifest.composeExtensions": "bot1,bot2",
        "manifest.staticTabs.contentUrl": "undefined,undefined",
        "manifest.configurableTabs.configurationUrl": "undefined",
        "manifest.webApplicationInfo.id": "web-app-id",
        "manifest.extensions": "false",
      })
    );
  });
});
