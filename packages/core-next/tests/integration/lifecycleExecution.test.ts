/**
 * Copyright (c) Microsoft Corporation.
 * Licensed under the MIT license.
 *
 * Integration test: YAML parse → driver registration → lifecycle execution
 * with real file-system drivers and env-var chaining between steps.
 */

import { expect } from "chai";
import { describe, it, beforeEach, afterEach } from "mocha";
import * as fs from "node:fs/promises";
import * as path from "node:path";
import * as os from "node:os";
import * as sinon from "sinon";

import { parseProjectYaml } from "../../src/lifecycle/parser";
import { executeLifecycle } from "../../src/lifecycle/executor";
import { registerBuiltinDrivers, builtinDrivers } from "../../src/drivers/builtin";
import { driverRegistry } from "../../src/drivers/registry";
import { createMockContext } from "../unit/testHelper";

/**
 * Creates a temporary directory for the test project.
 */
async function makeTmpDir(): Promise<string> {
  return fs.mkdtemp(path.join(os.tmpdir(), "atk-lifecycle-"));
}

describe("Integration: Lifecycle execution with real drivers", () => {
  let sandbox: sinon.SinonSandbox;

  beforeEach(() => {
    sandbox = sinon.createSandbox();
    registerBuiltinDrivers();
  });

  afterEach(() => {
    sandbox.restore();
  });

  it("should register all 22 builtin drivers", () => {
    expect(builtinDrivers).to.have.lengthOf(22);
    for (const d of builtinDrivers) {
      expect(driverRegistry.get(d.id)).to.not.be.undefined;
    }
  });

  it("should parse YAML and execute a provision lifecycle with env-var chaining", async () => {
    const tmpDir = await makeTmpDir();

    try {
      const yaml = `
version: "1.0.0"
environmentFolderPath: ./env

provision:
  - name: Write initial env
    uses: file/createOrUpdateEnvironmentFile
    with:
      target: env/.env.dev
      envs:
        APP_NAME: my-test-app
        API_ENDPOINT: https://api.example.com
    writeToEnvironmentFile:
      APP_NAME: APP_NAME
      API_ENDPOINT: API_ENDPOINT

  - name: Write config JSON from chained vars
    uses: file/createOrUpdateJsonFile
    with:
      target: config/settings.json
      content:
        appName: "\${{APP_NAME}}"
        endpoint: "\${{API_ENDPOINT}}"
        version: "1.0"
`;

      // Parse the YAML
      const parseResult = await parseProjectYaml(yaml);
      expect(parseResult.isOk()).to.be.true;
      const model = parseResult._unsafeUnwrap();

      expect(model.version).to.equal("1.0.0");
      expect(model.provision).to.have.lengthOf(2);

      // Execute the provision lifecycle
      const ctx = createMockContext({ projectPath: tmpDir });
      const envMap = new Map<string, string>();

      const execResult = await executeLifecycle(ctx, "provision", model.provision!, envMap);

      expect(execResult.isOk()).to.be.true;
      const result = execResult._unsafeUnwrap();
      expect(result.lifecycle).to.equal("provision");
      expect(result.steps).to.have.lengthOf(2);

      // Verify step 1: .env file was written
      const envContent = await fs.readFile(path.join(tmpDir, "env/.env.dev"), "utf-8");
      expect(envContent).to.include("APP_NAME=my-test-app");
      expect(envContent).to.include("API_ENDPOINT=https://api.example.com");

      // Verify env var chaining: outputs from step 1 are in envMap
      expect(envMap.get("APP_NAME")).to.equal("my-test-app");
      expect(envMap.get("API_ENDPOINT")).to.equal("https://api.example.com");

      // Verify step 2: JSON file was written with resolved placeholders
      const jsonContent = await fs.readFile(path.join(tmpDir, "config/settings.json"), "utf-8");
      const parsed = JSON.parse(jsonContent);
      expect(parsed.appName).to.equal("my-test-app");
      expect(parsed.endpoint).to.equal("https://api.example.com");
      expect(parsed.version).to.equal("1.0");
    } finally {
      await fs.rm(tmpDir, { recursive: true, force: true });
    }
  });

  it("should execute script driver and chain its outputs to subsequent steps", async () => {
    const tmpDir = await makeTmpDir();

    try {
      // Create a script that emits ::set-output
      const isWindows = os.platform() === "win32";
      const scriptFile = isWindows ? "gen.cmd" : "gen.sh";
      const scriptContent = isWindows
        ? "@echo off\r\necho ::set-output GENERATED_ID=gen-abc-123\r\necho ::set-output BUILD_NUM=42\r\n"
        : "#!/bin/sh\necho '::set-output GENERATED_ID=gen-abc-123'\necho '::set-output BUILD_NUM=42'\n";

      await fs.writeFile(path.join(tmpDir, scriptFile), scriptContent, { mode: 0o755 });

      const yaml = `
version: "1.0.0"

provision:
  - name: Generate IDs
    uses: script
    with:
      run: ${isWindows ? "gen.cmd" : "sh gen.sh"}
    writeToEnvironmentFile:
      GENERATED_ID: GENERATED_ID
      BUILD_NUM: BUILD_NUM

  - name: Write env with generated values
    uses: file/createOrUpdateEnvironmentFile
    with:
      target: output/.env.generated
      envs:
        RESOURCE_ID: "\${{GENERATED_ID}}"
        BUILD: "\${{BUILD_NUM}}"
`;

      const parseResult = await parseProjectYaml(yaml);
      expect(parseResult.isOk()).to.be.true;
      const model = parseResult._unsafeUnwrap();

      const ctx = createMockContext({ projectPath: tmpDir });
      const envMap = new Map<string, string>();

      const execResult = await executeLifecycle(ctx, "provision", model.provision!, envMap);
      expect(execResult.isOk()).to.be.true;

      const result = execResult._unsafeUnwrap();
      expect(result.steps).to.have.lengthOf(2);

      // Verify chained outputs
      expect(envMap.get("GENERATED_ID")).to.equal("gen-abc-123");
      expect(envMap.get("BUILD_NUM")).to.equal("42");

      // Verify the second step received resolved variables
      const envContent = await fs.readFile(path.join(tmpDir, "output/.env.generated"), "utf-8");
      expect(envContent).to.include("RESOURCE_ID=gen-abc-123");
      expect(envContent).to.include("BUILD=42");
    } finally {
      await fs.rm(tmpDir, { recursive: true, force: true });
    }
  });

  it("should return DriverNotFound for unknown driver ID", async () => {
    const ctx = createMockContext({ projectPath: "/tmp" });
    const envMap = new Map<string, string>();

    const steps = [
      {
        uses: "nonExistent/fakeDriver",
        with: { foo: "bar" },
      },
    ];

    const result = await executeLifecycle(ctx, "provision", steps, envMap);
    expect(result.isErr()).to.be.true;
    if (result.isErr()) {
      expect(result.error.code).to.equal("DriverNotFound");
    }
  });

  it("should handle multiple lifecycles from the same YAML", async () => {
    const tmpDir = await makeTmpDir();

    try {
      const yaml = `
version: "1.0.0"

registerApp:
  - name: Create env
    uses: file/createOrUpdateEnvironmentFile
    with:
      target: env/.env.local
      envs:
        TEAMS_APP_ID: placeholder-id
    writeToEnvironmentFile:
      TEAMS_APP_ID: TEAMS_APP_ID

deploy:
  - name: Write deploy config
    uses: file/createOrUpdateJsonFile
    with:
      target: deploy-config.json
      content:
        deployed: true
        timestamp: "2026-03-31"
`;

      const parseResult = await parseProjectYaml(yaml);
      expect(parseResult.isOk()).to.be.true;
      const model = parseResult._unsafeUnwrap();

      expect(model.registerApp).to.have.lengthOf(1);
      expect(model.deploy).to.have.lengthOf(1);

      // Execute registerApp
      const ctx = createMockContext({ projectPath: tmpDir });
      const envMap = new Map<string, string>();

      const regResult = await executeLifecycle(ctx, "registerApp", model.registerApp!, envMap);
      expect(regResult.isOk()).to.be.true;
      expect(envMap.get("TEAMS_APP_ID")).to.equal("placeholder-id");

      // Execute deploy (independent lifecycle, shared env)
      const deployResult = await executeLifecycle(ctx, "deploy", model.deploy!, envMap);
      expect(deployResult.isOk()).to.be.true;

      const jsonContent = await fs.readFile(path.join(tmpDir, "deploy-config.json"), "utf-8");
      const parsed = JSON.parse(jsonContent);
      expect(parsed.deployed).to.equal(true);
      expect(parsed.timestamp).to.equal("2026-03-31");
    } finally {
      await fs.rm(tmpDir, { recursive: true, force: true });
    }
  });

  it("should preserve unresolved placeholders with a warning", async () => {
    const tmpDir = await makeTmpDir();

    try {
      const yaml = `
version: "1.0.0"

provision:
  - name: Write with unresolved var
    uses: file/createOrUpdateEnvironmentFile
    with:
      target: env/.env.test
      envs:
        VALUE: "\${{UNDEFINED_VAR}}"
`;

      const parseResult = await parseProjectYaml(yaml);
      expect(parseResult.isOk()).to.be.true;
      const model = parseResult._unsafeUnwrap();

      const ctx = createMockContext({ projectPath: tmpDir });
      const envMap = new Map<string, string>();

      const execResult = await executeLifecycle(ctx, "provision", model.provision!, envMap);
      expect(execResult.isOk()).to.be.true;

      // Warning was logged about unresolved variable
      expect((ctx.logger.warning as sinon.SinonStub).called).to.be.true;
      const warningMsg = (ctx.logger.warning as sinon.SinonStub).firstCall.args[0] as string;
      expect(warningMsg).to.include("UNDEFINED_VAR");

      // Unresolved placeholder is kept as-is in the file
      const envContent = await fs.readFile(path.join(tmpDir, "env/.env.test"), "utf-8");
      expect(envContent).to.include("VALUE=${{UNDEFINED_VAR}}");
    } finally {
      await fs.rm(tmpDir, { recursive: true, force: true });
    }
  });
});
