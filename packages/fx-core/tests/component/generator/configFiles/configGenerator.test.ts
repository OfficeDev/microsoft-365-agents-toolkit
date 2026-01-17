// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { assert } from "chai";
import fs from "fs-extra";
import "mocha";
import os from "os";
import path from "path";
import { createSandbox, SinonSandbox } from "sinon";

import { featureFlagManager, FeatureFlags } from "../../../../src/common/featureFlags";
import { createContext, setTools } from "../../../../src/common/globalVars";
import { ProjectTypeProps, TelemetryEvent } from "../../../../src/common/telemetry";
import { configGenerator } from "../../../../src/component/generator/configFiles/configGenerator";
import * as folderUtils from "../../../../src/folder";
import { MockTools } from "../../../core/utils";

describe("ConfigGenerator", () => {
  let sandbox: SinonSandbox;
  let tempRoot: string;
  let templatesRoot: string;
  let destination: string;

  beforeEach(async () => {
    sandbox = createSandbox();
    tempRoot = path.join(os.tmpdir(), `cfg-gen-${Date.now()}`);
    templatesRoot = path.join(tempRoot, "templates");
    destination = path.join(tempRoot, "dest");
    await fs.ensureDir(templatesRoot);
    await fs.ensureDir(destination);

    // Point template lookup to our temp templates root
    sandbox.stub(folderUtils, "getTemplatesFolder").returns(templatesRoot);

    // Enable GenerateConfigFiles so settingsUtil reads newly generated m365agents.*.yml
    featureFlagManager.setBooleanValue(FeatureFlags.GenerateConfigFiles, true);

    // Provide tools/context
    const tools = new MockTools();
    setTools(tools);
  });

  afterEach(async () => {
    sandbox.restore();
    // Reset flag to default
    featureFlagManager.setBooleanValue(FeatureFlags.GenerateConfigFiles, false);
    await fs.remove(tempRoot);
  });

  it("generates and merges files for local-typescript", async () => {
    // Arrange templates for local/typescript
    const localTs = path.join(templatesRoot, "configs", "local", "typescript");
    await fs.ensureDir(path.join(localTs, ".vscode"));
    await fs.ensureDir(path.join(localTs, "env"));
    await fs.writeFile(path.join(localTs, "package.json"), '{"fromTemplate":true,"arr":[1,2]}');
    await fs.writeFile(path.join(localTs, ".vscode", "launch.json.tpl"), '{"version":"0.2.0"}');
    await fs.writeFile(
      path.join(localTs, ".vscode", "tasks.json.tpl"),
      '{"tasks":[{"label":"build"}]}'
    );
    await fs.writeFile(
      path.join(localTs, "m365agents.local.yml.tpl"),
      "version: 1.0.0\n# projectId will be added by settingsUtil\n"
    );
    await fs.writeFile(path.join(localTs, "env", ".env.local"), "# local env\n");

    // Pre-existing destination package.json to trigger JSON merge
    await fs.writeFile(
      path.join(destination, "package.json"),
      '{"name":"dest","existing":1,"arr":[1]}'
    );

    const tools = new MockTools();
    setTools(tools);
    const context = createContext();

    const telemetryStub = sandbox.stub(context.telemetryReporter, "sendTelemetryEvent");

    // Act
    const res = await configGenerator.run(
      context,
      destination,
      [{ name: "local", programmingLanguage: "typescript" }],
      { hasTab: true }
    );

    // Assert
    assert.isTrue(res.isOk(), "run should succeed");

    // JSON merged with de-dup array and preserved existing primitive
    const pkg = JSON.parse(await fs.readFile(path.join(destination, "package.json"), "utf8"));
    assert.equal(pkg.existing, 1);
    assert.deepEqual(pkg.arr, [1, 2]);

    // Rendered task copied and temp .rendered removed
    const tasksPath = path.join(destination, ".vscode", "tasks.json");
    assert.isTrue(await fs.pathExists(tasksPath), "tasks.json should exist");
    assert.isFalse(
      await fs.pathExists(tasksPath + ".rendered"),
      "temp rendered file should be cleaned"
    );

    // Generated yaml exists and should contain projectId after writeSettings
    const ymlPath = path.join(destination, "m365agents.local.yml");
    assert.isTrue(await fs.pathExists(ymlPath), "m365agents.local.yml should exist");
    const yml = await fs.readFile(ymlPath, "utf8");
    assert.match(yml, /projectId:\s*.+/);

    // Telemetry summary emitted with capabilities and success components
    const summaryCall = telemetryStub
      .getCalls()
      .find((c) => c.args[0] === TelemetryEvent.GenerateConfigSummary);
    assert.isOk(summaryCall, "should send GenerateConfigSummary telemetry");
    const props = summaryCall!.args[1] as Record<string, string>;
    assert.equal(props[ProjectTypeProps.TeamsManifestCapabilities], "Tab");
    assert.include(props.successComponents, "local-typescript");
    assert.equal(props.failedComponents, "");
    assert.isTrue((props.trackingId || "").length > 0);
  });

  it("continues on conflict and reports success/failed components", async () => {
    // Arrange templates for local/typescript
    const localTs = path.join(templatesRoot, "configs", "local", "typescript");
    await fs.ensureDir(path.join(localTs, ".vscode"));
    await fs.writeFile(path.join(localTs, "package.json"), '{"fromTemplate":true}');
    await fs.writeFile(path.join(localTs, ".vscode", "launch.json.tpl"), "{}");
    await fs.writeFile(path.join(localTs, ".vscode", "tasks.json.tpl"), "{}");
    await fs.writeFile(path.join(localTs, "m365agents.local.yml.tpl"), "version: 1.0.0\n");

    // Arrange templates for playground/typescript
    const pgTs = path.join(templatesRoot, "configs", "playground", "typescript");
    await fs.ensureDir(path.join(pgTs, ".vscode"));
    await fs.ensureDir(path.join(pgTs, "env"));
    await fs.writeFile(path.join(pgTs, "package.json"), '{"fromTemplate":true}');
    await fs.writeFile(path.join(pgTs, ".vscode", "launch.json"), "{}");
    await fs.writeFile(path.join(pgTs, ".vscode", "tasks.json"), "{}");
    await fs.writeFile(path.join(pgTs, "m365agents.playground.yml"), "version: 1.0.0\n");
    await fs.writeFile(path.join(pgTs, ".localConfigs.playground"), "{}");
    await fs.writeFile(path.join(pgTs, "env", ".env.playground"), "# playground env\n");
    await fs.writeFile(path.join(pgTs, "env", ".env.playground.user"), "# playground user env\n");

    // Create conflict file for local policy (env/.env.local has allowExistingFile=false)
    await fs.ensureDir(path.join(destination, "env"));
    await fs.writeFile(path.join(destination, "env", ".env.local"), "FOO=BAR\n");

    const tools = new MockTools();
    setTools(tools);
    const context = createContext();
    const telemetryStub = sandbox.stub(context.telemetryReporter, "sendTelemetryEvent");
    const showMsgStub = sandbox.stub(context.userInteraction, "showMessage");

    // Act
    const res = await configGenerator.run(
      context,
      destination,
      [
        { name: "local", programmingLanguage: "typescript" },
        { name: "playground", programmingLanguage: "typescript" },
      ],
      { hasBot: true }
    );

    // Assert: overall success because playground succeeded and provided yaml for settings
    assert.isTrue(res.isOk(), "run should succeed when at least one component succeeds");

    // Conflict warning shown
    assert.isTrue(showMsgStub.calledWith("warn"));

    // Playground yaml exists, local yaml not generated due to conflict
    assert.isTrue(await fs.pathExists(path.join(destination, "m365agents.playground.yml")));
    assert.isFalse(await fs.pathExists(path.join(destination, "m365agents.local.yml")));

    // Telemetry summary shows success/failed components and capabilities
    const summaryCall = telemetryStub
      .getCalls()
      .find((c) => c.args[0] === TelemetryEvent.GenerateConfigSummary);
    assert.isOk(summaryCall);
    const props = summaryCall!.args[1] as Record<string, string>;
    assert.include(props.successComponents, "playground-typescript");
    assert.include(props.failedComponents, "local-typescript");
    assert.equal(props[ProjectTypeProps.TeamsManifestCapabilities], "Bot");
  });

  it("returns error for unknown policy", async () => {
    const tools = new MockTools();
    setTools(tools);
    const context = createContext();

    const res = await configGenerator.run(
      context,
      destination,
      [{ name: "unknown", programmingLanguage: "typescript" }],
      {}
    );

    assert.isTrue(res.isErr(), "should return error for unknown policy");
    const errRes = res as any;
    const theErr = errRes.error as Error;
    assert.equal((theErr as any).name, "UnknownPolicyError");
    assert.equal((theErr as any).source, "ConfigGenerator");
  });
});
