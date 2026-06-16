// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import {
  AppManifestUtils,
  Inputs,
  ok,
  Platform,
  TeamsManifest,
  UserInteraction,
} from "@microsoft/teamsfx-api";
import { assert } from "chai";
import { createSandbox, SinonSandbox, SinonStub } from "sinon";
import { setTools, TOOLS } from "../../src/common/globalVars";
import { configGenerator } from "../../src/component/generator/configFiles/configGenerator";
import { generateConfigFiles } from "../../src/core/generateConfigFiles";
import { MockTools } from "./utils";

describe("generateConfigFiles", () => {
  let sandbox: SinonSandbox;
  let runStub: SinonStub;
  let readManifestStub: SinonStub;
  let showMessageStub: SinonStub;
  let mockTools: MockTools;
  const manifestPath = "appPackage/manifest.json";
  const projectPath = "/tmp/project";
  const programmingLanguage = "typescript";
  const originalTools = TOOLS;

  beforeEach(() => {
    sandbox = createSandbox();
    type RunResult = Awaited<ReturnType<typeof configGenerator.run>>;
    readManifestStub = sandbox.stub(AppManifestUtils, "readTeamsManifest");
    runStub = sandbox.stub(configGenerator, "run").resolves(ok({}) as unknown as RunResult);
    showMessageStub = sandbox.stub();
    mockTools = new MockTools();
    mockTools.ui = { showMessage: showMessageStub } as unknown as UserInteraction;
    setTools(mockTools);
  });

  afterEach(() => {
    sandbox.restore();
    setTools(originalTools);
  });

  const createInputs = (overrides: Record<string, unknown>): Inputs => {
    return {
      platform: Platform.CLI,
      projectPath,
      ...overrides,
    } as Inputs;
  };

  it("generates playground and local configs when bot present", async () => {
    const manifest = {
      name: { short: "MyApp" },
      bots: [{ botId: "bot-id", scopes: ["team"] }],
      staticTabs: [
        {
          entityId: "entity",
          name: "Tab",
          contentUrl: "https://example.com",
          websiteUrl: "https://example.com",
          scopes: ["team"],
        },
      ],
    } as unknown as TeamsManifest;
    readManifestStub.resolves(manifest);

    await generateConfigFiles(
      createInputs({
        "manifest-file": manifestPath,
        "include-playground": true,
        "include-local": true,
        "include-remote": false,
        "programming-language": programmingLanguage,
      })
    );

    assert.isTrue(runStub.calledOnce);
    const args = runStub.firstCall.args;
    const components = args[2];
    assert.deepEqual(components, [
      { name: "playground", programmingLanguage },
      { name: "local", programmingLanguage },
    ]);
    const features = args[3];
    assert.deepInclude(features, { hasBot: true, hasTab: true });
    assert.equal(features.appName, "MyApp");
    assert.isTrue(showMessageStub.notCalled);
  });

  it("skips playground and warns when bot missing", async () => {
    const manifest = {
      name: { short: "NoBotApp" },
      staticTabs: [],
      bots: [],
    } as unknown as TeamsManifest;
    readManifestStub.resolves(manifest);

    await generateConfigFiles(
      createInputs({
        "manifest-file": manifestPath,
        "include-playground": true,
        "include-local": true,
        "include-remote": false,
        "programming-language": programmingLanguage,
      })
    );

    assert.isTrue(runStub.calledOnce);
    const components = runStub.firstCall.args[2];
    assert.deepEqual(components, [{ name: "local", programmingLanguage }]);
    assert.isTrue(showMessageStub.calledOnce);
    assert.equal(showMessageStub.firstCall.args[0], "warn");
  });

  it("generates only playground when include-local is false", async () => {
    const manifest = {
      name: { short: "PlaygroundOnly" },
      bots: [{ botId: "bot-id", scopes: ["team"] }],
    } as unknown as TeamsManifest;
    readManifestStub.resolves(manifest);

    await generateConfigFiles(
      createInputs({
        "manifest-file": manifestPath,
        "include-playground": true,
        "include-local": false,
        "include-remote": false,
        "programming-language": programmingLanguage,
      })
    );

    const components = runStub.firstCall.args[2];
    assert.deepEqual(components, [{ name: "playground", programmingLanguage }]);
  });

  it("propagates copilot feature flag", async () => {
    const manifest = {
      name: { short: "CopilotApp" },
      bots: [{ botId: "bot-id", scopes: ["team"] }],
      copilotAgents: [{ id: "agent1" }],
    } as unknown as TeamsManifest;
    readManifestStub.resolves(manifest);

    await generateConfigFiles(
      createInputs({
        "manifest-file": manifestPath,
        "include-playground": true,
        "include-local": false,
        "include-remote": false,
        "programming-language": programmingLanguage,
      })
    );

    const features = runStub.firstCall.args[3] as Record<string, unknown>;
    assert.isTrue(features.hasCopilot as boolean);
  });

  it("includes remote config when include-remote is true", async () => {
    const manifest = {
      name: { short: "RemoteApp" },
      staticTabs: [],
      bots: [],
    } as unknown as TeamsManifest;
    readManifestStub.resolves(manifest);

    await generateConfigFiles(
      createInputs({
        "manifest-file": manifestPath,
        "include-playground": false,
        "include-local": false,
        "include-remote": true,
        "programming-language": programmingLanguage,
      })
    );

    assert.isTrue(runStub.calledOnce);
    const components = runStub.firstCall.args[2];
    assert.deepEqual(components, [{ name: "remote", programmingLanguage }]);
  });

  it("generates all configs including remote", async () => {
    const manifest = {
      name: { short: "FullApp" },
      bots: [{ botId: "bot-id", scopes: ["team"] }],
      staticTabs: [
        {
          entityId: "entity",
          name: "Tab",
          contentUrl: "https://example.com",
          websiteUrl: "https://example.com",
          scopes: ["team"],
        },
      ],
    } as unknown as TeamsManifest;
    readManifestStub.resolves(manifest);

    await generateConfigFiles(
      createInputs({
        "manifest-file": manifestPath,
        "include-playground": true,
        "include-local": true,
        "include-remote": true,
        "programming-language": programmingLanguage,
      })
    );

    assert.isTrue(runStub.calledOnce);
    const components = runStub.firstCall.args[2];
    assert.deepEqual(components, [
      { name: "playground", programmingLanguage },
      { name: "local", programmingLanguage },
      { name: "remote", programmingLanguage },
    ]);
  });

  it("excludes remote config when include-remote is false", async () => {
    const manifest = {
      name: { short: "LocalOnlyApp" },
      bots: [{ botId: "bot-id", scopes: ["team"] }],
    } as unknown as TeamsManifest;
    readManifestStub.resolves(manifest);

    await generateConfigFiles(
      createInputs({
        "manifest-file": manifestPath,
        "include-playground": true,
        "include-local": true,
        "include-remote": false,
        "programming-language": programmingLanguage,
      })
    );

    const components = runStub.firstCall.args[2];
    assert.deepEqual(components, [
      { name: "playground", programmingLanguage },
      { name: "local", programmingLanguage },
    ]);
    assert.isFalse(components.map((c: any) => c.name).includes("remote"));
  });
});
