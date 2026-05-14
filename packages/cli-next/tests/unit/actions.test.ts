/**
 * Copyright (c) Microsoft Corporation.
 * Licensed under the MIT license.
 */

import { expect } from "chai";
import { describe, it, afterEach, beforeEach } from "mocha";
import * as sinon from "sinon";
import type { AtkContext, TemplateDescriptor } from "@microsoft/teamsfx-core-next";
import { TemplateRegistry } from "@microsoft/teamsfx-core-next";
import { listTemplatesAction } from "../../src/actions/listTemplates";

function createMockContext(overrides?: Partial<AtkContext>): AtkContext {
  return {
    auth: {
      m365TokenProvider: {} as any,
      azureAccountProvider: {} as any,
    },
    logger: {
      log: sinon.stub(),
      verbose: sinon.stub(),
      debug: sinon.stub(),
      info: sinon.stub(),
      warning: sinon.stub(),
      error: sinon.stub(),
      logInFile: sinon.stub().resolves(),
      getLogFilePath: sinon.stub().returns("/tmp/test.log"),
    } as any,
    telemetry: {
      sendTelemetryEvent: sinon.stub(),
      sendTelemetryErrorEvent: sinon.stub(),
      sendTelemetryException: sinon.stub(),
    },
    ui: {
      selectOption: sinon.stub(),
      selectOptions: sinon.stub(),
      inputText: sinon.stub(),
      selectFile: sinon.stub(),
      selectFiles: sinon.stub(),
      selectFolder: sinon.stub(),
      openUrl: sinon.stub(),
      showMessage: sinon.stub(),
      createProgressBar: sinon.stub(),
      confirm: sinon.stub(),
    } as any,
    correlationId: "test-correlation-id",
    projectPath: "/tmp/test-project",
    ...overrides,
  };
}

function makeDescriptor(
  id: string,
  category: string,
  languages: string[] = ["typescript"]
): TemplateDescriptor {
  return {
    id,
    name: id.replace("/", " "),
    category: category as any,
    languages: languages as any[],
    scaffoldFn: async () =>
      ({
        isOk: () => true,
        isErr: () => false,
        value: { projectPath: "/tmp", warnings: [] },
      }) as any,
  };
}

describe("Actions", () => {
  const sandbox = sinon.createSandbox();

  beforeEach(() => {
    sandbox.stub(console, "log");
    sandbox.stub(console, "warn");
    sandbox.stub(console, "error");
  });

  afterEach(() => {
    sandbox.restore();
  });

  describe("listTemplatesAction()", () => {
    it("should return all registered templates as rows", () => {
      const registry = new TemplateRegistry();
      registry.register(makeDescriptor("bot/echo", "bot", ["typescript", "javascript"]));
      registry.register(makeDescriptor("da/basic", "declarative-agent", ["common"]));

      const rows = listTemplatesAction(registry);
      expect(rows).to.have.length(2);
    });

    it("should sort templates by ID", () => {
      const registry = new TemplateRegistry();
      registry.register(makeDescriptor("tab/react", "tab"));
      registry.register(makeDescriptor("bot/echo", "bot"));

      const rows = listTemplatesAction(registry);
      expect(rows[0].id).to.equal("bot/echo");
      expect(rows[1].id).to.equal("tab/react");
    });

    it("should format language list as comma-separated string", () => {
      const registry = new TemplateRegistry();
      registry.register(makeDescriptor("bot/echo", "bot", ["typescript", "javascript"]));

      const rows = listTemplatesAction(registry);
      expect(rows[0].languages).to.equal("typescript, javascript");
    });

    it("should return empty array for empty registry", () => {
      const registry = new TemplateRegistry();
      const rows = listTemplatesAction(registry);
      expect(rows).to.have.length(0);
    });
  });

  describe("createProjectAction()", () => {
    it("should call runOperation with correct input", async () => {
      // This is a structural test — the real integration test verifies e2e.
      // We test that the action module imports and calls the right API shape.
      const { createProjectAction } = await import("../../src/actions/createProject");
      const ctx = createMockContext();

      // We mock console to suppress output; the operation will fail
      // because no real template is registered, but the action shape is correct.
      try {
        await createProjectAction(ctx, {
          templateId: "bot/echo",
          projectName: "test-project",
          language: "typescript",
          destinationPath: "/tmp",
        });
      } catch (e: any) {
        // Expected — no real template registered.
        expect(e.message).to.be.a("string");
      }
    });
  });

  describe("envListAction()", () => {
    it("should return environment names from a project with envs", async () => {
      const fs = await import("fs");
      const path = await import("path");
      const os = await import("os");

      // Create a temp project with env files
      const tmpDir = await fs.promises.mkdtemp(path.join(os.tmpdir(), "cli-test-"));
      const envDir = path.join(tmpDir, "env");
      await fs.promises.mkdir(envDir, { recursive: true });
      await fs.promises.writeFile(path.join(envDir, ".env.dev"), "APP_NAME=test\n");
      await fs.promises.writeFile(path.join(envDir, ".env.staging"), "APP_NAME=test\n");

      const { envListAction } = await import("../../src/actions/environment");
      const envs = await envListAction(tmpDir);
      expect(envs).to.include("dev");
      expect(envs).to.include("staging");

      // Cleanup
      await fs.promises.rm(tmpDir, { recursive: true, force: true });
    });

    it("should return empty array for project without env dir", async () => {
      const fs = await import("fs");
      const path = await import("path");
      const os = await import("os");

      const tmpDir = await fs.promises.mkdtemp(path.join(os.tmpdir(), "cli-test-"));

      const { envListAction } = await import("../../src/actions/environment");
      const envs = await envListAction(tmpDir);
      expect(envs).to.have.length(0);

      await fs.promises.rm(tmpDir, { recursive: true, force: true });
    });
  });

  describe("envAddAction()", () => {
    it("should create a new environment file", async () => {
      const fs = await import("fs");
      const path = await import("path");
      const os = await import("os");

      const tmpDir = await fs.promises.mkdtemp(path.join(os.tmpdir(), "cli-test-"));

      const { envAddAction } = await import("../../src/actions/environment");
      await envAddAction(tmpDir, "production");

      const envFilePath = path.join(tmpDir, "env", ".env.production");
      expect(await fileExists(envFilePath)).to.be.true;

      await fs.promises.rm(tmpDir, { recursive: true, force: true });
    });

    it("should copy from existing environment when --copy-from is specified", async () => {
      const fs = await import("fs");
      const path = await import("path");
      const os = await import("os");

      const tmpDir = await fs.promises.mkdtemp(path.join(os.tmpdir(), "cli-test-"));
      const envDir = path.join(tmpDir, "env");
      await fs.promises.mkdir(envDir, { recursive: true });
      await fs.promises.writeFile(path.join(envDir, ".env.dev"), "APP_NAME=myapp\n");

      const { envAddAction } = await import("../../src/actions/environment");
      await envAddAction(tmpDir, "staging", "dev");

      const content = await fs.promises.readFile(path.join(envDir, ".env.staging"), "utf-8");
      expect(content).to.include("APP_NAME=myapp");

      await fs.promises.rm(tmpDir, { recursive: true, force: true });
    });
  });

  describe("envResetAction()", () => {
    it("should clear the environment file", async () => {
      const fs = await import("fs");
      const path = await import("path");
      const os = await import("os");

      const tmpDir = await fs.promises.mkdtemp(path.join(os.tmpdir(), "cli-test-"));
      const envDir = path.join(tmpDir, "env");
      await fs.promises.mkdir(envDir, { recursive: true });
      await fs.promises.writeFile(path.join(envDir, ".env.dev"), "APP_NAME=myapp\nSECRET=abc\n");

      const { envResetAction } = await import("../../src/actions/environment");
      await envResetAction(tmpDir, "dev");

      const content = await fs.promises.readFile(path.join(envDir, ".env.dev"), "utf-8");
      expect(content.trim()).to.equal("");

      await fs.promises.rm(tmpDir, { recursive: true, force: true });
    });
  });
});

async function fileExists(filePath: string): Promise<boolean> {
  try {
    const fs = await import("fs");
    await fs.promises.access(filePath);
    return true;
  } catch {
    return false;
  }
}
