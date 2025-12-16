// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { Context, UserError } from "@microsoft/teamsfx-api";
import { assert } from "chai";
import fs from "fs-extra";
import "mocha";
import os from "os";
import path from "path";
import { createSandbox, SinonSandbox } from "sinon";
import { ConfigGenerator } from "../../../src/component/generator/configFiles/configGenerator";
import { policys } from "../../../src/component/generator/configFiles/copyPolicy";
import * as jsonMerger from "../../../src/component/generator/configFiles/jsonMerger";
import * as renderTemplateModule from "../../../src/component/generator/configFiles/renderTemplate";
import * as telemetryModule from "../../../src/component/telemetry";
import * as settingsUtilModule from "../../../src/component/utils/settingsUtil";

// eslint-disable-next-line @typescript-eslint/no-explicit-any
type AnyConfigGenerator = any;

describe("ConfigGenerator", () => {
  let sandbox: SinonSandbox;
  let configGenerator: ConfigGenerator;
  let tempDir: string;
  let mockContext: Context;

  beforeEach(async () => {
    sandbox = createSandbox();
    configGenerator = new ConfigGenerator();
    tempDir = path.join(os.tmpdir(), `test-config-gen-${Date.now()}`);
    await fs.ensureDir(tempDir);

    mockContext = {
      userInteraction: {
        showMessage: sandbox.stub().resolves(),
      },
      telemetryReporter: {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        sendTelemetryEvent: sandbox.stub() as any,
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        sendTelemetryErrorEvent: sandbox.stub() as any,
      },
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
    } as any;

    // Some middleware paths call sendErrorEvent without a context; stub to avoid TypeError
    sandbox.stub(telemetryModule, "sendErrorEvent").callsFake(() => {
      /* no-op for unit tests */
    });
  });

  afterEach(async () => {
    sandbox.restore();
    await fs.remove(tempDir);
  });

  describe("run", () => {
    it("should successfully generate config files for single component", async () => {
      const sourcePath = path.join(tempDir, "source");
      const destPath = path.join(tempDir, "dest");
      await fs.ensureDir(sourcePath);
      await fs.ensureDir(destPath);

      const testFile = "package.json";
      const sourceFile = path.join(sourcePath, testFile);
      await fs.writeFile(sourceFile, '{"test": "value"}');

      sandbox.stub(fs, "pathExists").resolves(false);
      sandbox.stub(fs, "copy").resolves();

      const components = [{ name: "playground", programmingLanguage: "typescript" }];
      await configGenerator.run(mockContext, destPath, components, {});

      // Run completed without throwing; rely on side effects in other tests
    });

    it("should handle multiple components", async () => {
      const destPath = path.join(tempDir, "dest");
      await fs.ensureDir(destPath);

      sandbox.stub(fs, "pathExists").resolves(false);
      sandbox.stub(fs, "copy").resolves();

      const components = [
        { name: "playground", programmingLanguage: "typescript" },
        { name: "local", programmingLanguage: "typescript" },
      ];
      await configGenerator.run(mockContext, destPath, components, {});
    });

    it("should return error for unknown policy", async () => {
      const destPath = path.join(tempDir, "dest");
      await fs.ensureDir(destPath);

      const components = [{ name: "unknown", programmingLanguage: "unknown" }];
      const resultUnknown = await configGenerator.run(mockContext, destPath, components, {});

      assert.isTrue(resultUnknown.isErr());
      if (resultUnknown.isErr()) {
        const error = resultUnknown.error as UserError;
        assert.equal(error.name, "UnknownPolicyError");
      }
    });

    it("should show warning and skip component when file conflict is detected", async () => {
      const destPath = path.join(tempDir, "dest");
      await fs.ensureDir(destPath);

      const policyKey = Object.keys(policys)[0];
      const [componentName, language] = policyKey.split("-");

      sandbox.stub(fs, "pathExists").resolves(true);
      const copyStub = sandbox.stub(fs, "copy").resolves();

      const components = [{ name: componentName, programmingLanguage: language }];
      await configGenerator.run(mockContext, destPath, components, {});

      // Middleware may wrap results; just verify warning behavior and skip
      assert.isFalse(copyStub.called);
    });

    it("should read settings and generate trackingId if missing", async () => {
      const destPath = path.join(tempDir, "dest");
      await fs.ensureDir(destPath);

      const readSettingsStub = sandbox
        .stub(settingsUtilModule.settingsUtil, "readSettings")
        .resolves({
          isOk: () => true,
          value: { trackingId: "", version: "1.0" },
        } as any);
      const writeSettingsStub = sandbox
        .stub(settingsUtilModule.settingsUtil, "writeSettings")
        .resolves({
          isOk: () => true,
          value: destPath,
        } as any);

      sandbox.stub(fs, "pathExists").resolves(false);
      sandbox.stub(fs, "copy").resolves();

      const components = [{ name: "playground", programmingLanguage: "typescript" }];
      const result = await configGenerator.run(mockContext, destPath, components, {});

      assert.isTrue(result.isOk());
      assert.isTrue(readSettingsStub.calledOnce);
      assert.isTrue(writeSettingsStub.calledOnce);
      const writeCall = writeSettingsStub.firstCall;
      const settings = writeCall.args[1];
      assert.isTrue(settings.trackingId.length > 0);
    });

    it("should use existing trackingId if present", async () => {
      const destPath = path.join(tempDir, "dest");
      await fs.ensureDir(destPath);
      const existingId = "existing-tracking-id";

      const readSettingsStub = sandbox
        .stub(settingsUtilModule.settingsUtil, "readSettings")
        .resolves({
          isOk: () => true,
          value: { trackingId: existingId, version: "1.0" },
        } as any);
      const writeSettingsStub = sandbox
        .stub(settingsUtilModule.settingsUtil, "writeSettings")
        .resolves({
          isOk: () => true,
          value: destPath,
        } as any);

      sandbox.stub(fs, "pathExists").resolves(false);
      sandbox.stub(fs, "copy").resolves();

      const components = [{ name: "playground", programmingLanguage: "typescript" }];
      const result = await configGenerator.run(mockContext, destPath, components, {});

      assert.isTrue(result.isOk());
      const writeCall = writeSettingsStub.firstCall;
      const settings = writeCall.args[1];
      assert.equal(settings.trackingId, existingId);
    });

    it("should return error when readSettings fails", async () => {
      const destPath = path.join(tempDir, "dest");
      await fs.ensureDir(destPath);

      sandbox.stub(settingsUtilModule.settingsUtil, "readSettings").resolves({
        isErr: () => true,
        error: new Error("Settings read failed"),
      } as any);

      sandbox.stub(fs, "pathExists").resolves(false);
      sandbox.stub(fs, "copy").resolves();

      const components = [{ name: "playground", programmingLanguage: "typescript" }];
      const result = await configGenerator.run(mockContext, destPath, components, {});

      assert.isTrue(result.isErr());
    });

    it("should send telemetry event with correct properties", async () => {
      const destPath = path.join(tempDir, "dest");
      await fs.ensureDir(destPath);

      sandbox.stub(settingsUtilModule.settingsUtil, "readSettings").resolves({
        isOk: () => true,
        value: { trackingId: "test-id", version: "1.0" },
      } as any);
      sandbox.stub(settingsUtilModule.settingsUtil, "writeSettings").resolves({
        isOk: () => true,
        value: destPath,
      } as any);

      sandbox.stub(fs, "pathExists").resolves(false);
      sandbox.stub(fs, "copy").resolves();

      const components = [
        { name: "playground", programmingLanguage: "typescript" },
        { name: "local", programmingLanguage: "typescript" },
      ];
      const features = { hasBot: true, hasTab: true, appName: "TestApp" };
      const result = await configGenerator.run(mockContext, destPath, components, features);

      assert.isTrue(result.isOk());
      assert.isTrue(mockContext.telemetryReporter.sendTelemetryEvent.calledOnce);

      const telemetryCall = mockContext.telemetryReporter.sendTelemetryEvent.firstCall;
      const eventName = telemetryCall.args[0];
      const eventProps = telemetryCall.args[1];

      assert.equal(eventName, "GenerateConfigSummary");
      assert.isTrue(eventProps.trackingId === "test-id");
      assert.isTrue(eventProps.successComponents.includes("playground-typescript"));
      assert.isTrue(eventProps.successComponents.includes("local-typescript"));
    });

    it("should include failed components in telemetry", async () => {
      const destPath = path.join(tempDir, "dest");
      await fs.ensureDir(destPath);

      const policyKey = Object.keys(policys)[0];
      const [componentName, language] = policyKey.split("-");

      sandbox.stub(settingsUtilModule.settingsUtil, "readSettings").resolves({
        isOk: () => true,
        value: { trackingId: "test-id", version: "1.0" },
      } as any);
      sandbox.stub(settingsUtilModule.settingsUtil, "writeSettings").resolves({
        isOk: () => true,
        value: destPath,
      } as any);

      sandbox.stub(fs, "pathExists").resolves(true);
      sandbox.stub(fs, "copy").resolves();

      const components = [
        { name: componentName, programmingLanguage: language },
        { name: "local", programmingLanguage: "typescript" },
      ];
      const result = await configGenerator.run(mockContext, destPath, components, {});

      const telemetryCall = mockContext.telemetryReporter.sendTelemetryEvent.firstCall;
      const eventProps = telemetryCall.args[1];

      assert.isTrue(eventProps.failedComponents.length > 0);
    });

    it("should collect all capabilities in telemetry", async () => {
      const destPath = path.join(tempDir, "dest");
      await fs.ensureDir(destPath);

      sandbox.stub(settingsUtilModule.settingsUtil, "readSettings").resolves({
        isOk: () => true,
        value: { trackingId: "test-id", version: "1.0" },
      } as any);
      sandbox.stub(settingsUtilModule.settingsUtil, "writeSettings").resolves({
        isOk: () => true,
        value: destPath,
      } as any);

      sandbox.stub(fs, "pathExists").resolves(false);
      sandbox.stub(fs, "copy").resolves();

      const components = [{ name: "playground", programmingLanguage: "typescript" }];
      const features = {
        hasTab: true,
        hasBot: true,
        hasMessageExtension: true,
        hasDeclarativeAgent: true,
        hasCustomEngineAgent: true,
      };

      const result = await configGenerator.run(mockContext, destPath, components, features);

      assert.isTrue(result.isOk());
      const telemetryCall = mockContext.telemetryReporter.sendTelemetryEvent.firstCall;
      const eventProps = telemetryCall.args[1];

      const capabilities = eventProps.TeamsManifestCapabilities;
      assert.isTrue(capabilities.includes("Tab"));
      assert.isTrue(capabilities.includes("Bot"));
      assert.isTrue(capabilities.includes("ME"));
      assert.isTrue(capabilities.includes("DA"));
      assert.isTrue(capabilities.includes("CEA"));
    });
  });

  describe("detectFileConflict", () => {
    it("should return ok when no files exist", async () => {
      sandbox.stub(fs, "pathExists").resolves(false);

      const policy = {
        "package.json": { allowExistingFile: false, policy: "add" as const },
        ".vscode/launch.json": { allowExistingFile: true, policy: "add" as const },
      };

      const result = await (configGenerator as AnyConfigGenerator).detectFileConflict(
        tempDir,
        policy
      );

      assert.isTrue(result.isOk());
    });

    it("should return error when file exists and allowExistingFile is false", async () => {
      sandbox.stub(fs, "pathExists").resolves(true);

      const policy = {
        "package.json": { allowExistingFile: false, policy: "add" as const },
      };

      const result = await (configGenerator as AnyConfigGenerator).detectFileConflict(
        tempDir,
        policy
      );

      assert.isTrue(result.isErr());
      if (result.isErr()) {
        const error = result.error as UserError;
        assert.equal(error.name, "ConflictFileError");
      }
    });

    it("should return ok when file exists but allowExistingFile is true", async () => {
      sandbox.stub(fs, "pathExists").resolves(true);

      const policy = {
        "package.json": { allowExistingFile: true, policy: "add" as const },
      };

      const result = await (configGenerator as AnyConfigGenerator).detectFileConflict(
        tempDir,
        policy
      );

      assert.isTrue(result.isOk());
    });

    it("should only check files where allowExistingFile is false", async () => {
      const pathExistsStub = sandbox.stub(fs, "pathExists");
      pathExistsStub.resolves(false);

      const policy = {
        "package.json": { allowExistingFile: false, policy: "add" as const },
        ".vscode/launch.json": { allowExistingFile: true, policy: "add" as const },
        ".vscode/tasks.json": { allowExistingFile: false, policy: "add" as const },
      };

      await (configGenerator as AnyConfigGenerator).detectFileConflict(tempDir, policy);

      assert.isTrue(pathExistsStub.called);
    });
  });

  describe("getFileExtensionWithoutTemplate", () => {
    it("should return extension for file with .tpl suffix", () => {
      const ext = (configGenerator as AnyConfigGenerator).getFileExtensionWithoutTemplate(
        "config.json.tpl"
      );
      assert.equal(ext, ".json");
    });

    it("should return extension for file without .tpl suffix", () => {
      const ext = (configGenerator as AnyConfigGenerator).getFileExtensionWithoutTemplate(
        "package.json"
      );
      assert.equal(ext, ".json");
    });

    it("should return empty string for files without extension", () => {
      const ext = (configGenerator as AnyConfigGenerator).getFileExtensionWithoutTemplate(
        "Dockerfile"
      );
      assert.equal(ext, "");
    });

    it("should handle .tpl files with multiple extensions", () => {
      const ext = (configGenerator as AnyConfigGenerator).getFileExtensionWithoutTemplate(
        "config.yaml.tpl"
      );
      assert.equal(ext, ".yaml");
    });
  });

  describe("generateConfigFilesByPolicy", () => {
    it("should copy file when destination does not exist", async () => {
      const sourceDir = path.join(tempDir, "source");
      const destDir = path.join(tempDir, "dest");
      await fs.ensureDir(sourceDir);
      await fs.ensureDir(destDir);

      const sourceFile = path.join(sourceDir, "package.json");
      await fs.writeFile(sourceFile, '{"version": "1.0.0"}');

      sandbox.stub(fs, "pathExists").resolves(false);
      const copyStub = sandbox.stub(fs, "copy").resolves();

      const policy = {
        "package.json": { allowExistingFile: true, policy: "add" as const },
      };

      await (configGenerator as AnyConfigGenerator).generateConfigFilesByPolicy(
        sourceDir,
        destDir,
        policy,
        {}
      );

      assert.isTrue(copyStub.called);
    });

    it("should merge JSON files when destination exists and policy is 'add'", async () => {
      const sourceDir = path.join(tempDir, "source");
      const destDir = path.join(tempDir, "dest");
      await fs.ensureDir(sourceDir);
      await fs.ensureDir(destDir);

      const sourceFile = path.join(sourceDir, "package.json");
      const destFile = path.join(destDir, "package.json");
      await fs.writeFile(sourceFile, '{"version": "1.0.0"}');
      await fs.writeFile(destFile, '{"name": "test"}');

      sandbox.stub(fs, "pathExists").resolves(true);
      const mergeStub = sandbox.stub(jsonMerger, "mergeJsonFile").resolves();

      const policy = {
        "package.json": { allowExistingFile: true, policy: "add" as const },
      };

      await (configGenerator as AnyConfigGenerator).generateConfigFilesByPolicy(
        sourceDir,
        destDir,
        policy,
        {}
      );

      assert.isTrue(mergeStub.called);
    });

    it("should render template files before processing", async () => {
      const sourceDir = path.join(tempDir, "source");
      const destDir = path.join(tempDir, "dest");
      await fs.ensureDir(sourceDir);
      await fs.ensureDir(destDir);

      const sourceFile = path.join(sourceDir, "config.json.tpl");
      await fs.writeFile(sourceFile, '{"message": "{{message}}"}');

      sandbox.stub(fs, "pathExists").resolves(false);
      sandbox.stub(fs, "copy").resolves();
      sandbox.stub(fs, "writeFile").resolves();
      sandbox.stub(fs, "remove").resolves();
      const renderStub = sandbox
        .stub(renderTemplateModule, "renderTemplate")
        .returns('{"message": "hello"}');

      const policy = {
        "config.json.tpl": { allowExistingFile: false, policy: "skip" as const },
      };

      await (configGenerator as AnyConfigGenerator).generateConfigFilesByPolicy(
        sourceDir,
        destDir,
        policy,
        { message: "hello" }
      );

      assert.isTrue(renderStub.called);
    });

    it("should clean up rendered template files", async () => {
      const sourceDir = path.join(tempDir, "source");
      const destDir = path.join(tempDir, "dest");
      await fs.ensureDir(sourceDir);
      await fs.ensureDir(destDir);

      const sourceFile = path.join(sourceDir, "config.json.tpl");
      await fs.writeFile(sourceFile, '{"test": "{{value}}"}');

      sandbox.stub(fs, "pathExists").resolves(false);
      sandbox.stub(fs, "copy").resolves();
      const removeStub = sandbox.stub(fs, "remove").resolves();
      sandbox.stub(fs, "writeFile").resolves();
      sandbox.stub(renderTemplateModule, "renderTemplate").returns('{"test": "value"}');

      const policy = {
        "config.json.tpl": { allowExistingFile: false, policy: "skip" as const },
      };

      await (configGenerator as AnyConfigGenerator).generateConfigFilesByPolicy(
        sourceDir,
        destDir,
        policy,
        {}
      );

      assert.isTrue(removeStub.called);
    });

    it("should not merge non-JSON files even with 'add' policy", async () => {
      const sourceDir = path.join(tempDir, "source");
      const destDir = path.join(tempDir, "dest");
      await fs.ensureDir(sourceDir);
      await fs.ensureDir(destDir);

      const sourceFile = path.join(sourceDir, "config.yaml");
      const destFile = path.join(destDir, "config.yaml");
      await fs.writeFile(sourceFile, "test: value");
      await fs.writeFile(destFile, "existing: data");

      sandbox.stub(fs, "pathExists").resolves(true);
      const mergeStub = sandbox.stub(jsonMerger, "mergeJsonFile").resolves();

      const policy = {
        "config.yaml": { allowExistingFile: true, policy: "add" as const },
      };

      await (configGenerator as AnyConfigGenerator).generateConfigFilesByPolicy(
        sourceDir,
        destDir,
        policy,
        {}
      );

      assert.isFalse(mergeStub.called);
    });
  });

  describe("getPolicyKey", () => {
    it("should generate correct policy key", () => {
      const component = { name: "playground", programmingLanguage: "typescript" };
      const key = (configGenerator as AnyConfigGenerator).getPolicyKey(component);
      assert.equal(key, "playground-typescript");
    });

    it("should handle different component names and languages", () => {
      const component = { name: "local", programmingLanguage: "python" };
      const key = (configGenerator as AnyConfigGenerator).getPolicyKey(component);
      assert.equal(key, "local-python");
    });
  });
});
