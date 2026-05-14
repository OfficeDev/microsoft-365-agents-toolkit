/**
 * Copyright (c) Microsoft Corporation.
 * Licensed under the MIT license.
 */

import { expect } from "chai";
import { describe, it, beforeEach, afterEach } from "mocha";
import * as fs from "fs/promises";
import * as path from "path";
import * as os from "os";
import {
  copyExistMetaOSProject,
  extendToDA,
  unifyProjectID,
} from "../../../src/templates/helpers/metaOSHelper";

describe("MetaOSHelper", () => {
  let tmpDir: string;

  beforeEach(async () => {
    tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), "metaos-test-"));
  });

  afterEach(async () => {
    await fs.rm(tmpDir, { recursive: true, force: true });
  });

  describe("copyExistMetaOSProject", () => {
    it("should copy files excluding node_modules and env", async () => {
      const src = path.join(tmpDir, "source");
      const dest = path.join(tmpDir, "dest");

      // set up source project
      await fs.mkdir(path.join(src, "src"), { recursive: true });
      await fs.mkdir(path.join(src, "node_modules", "dep"), { recursive: true });
      await fs.mkdir(path.join(src, "env"), { recursive: true });
      await fs.writeFile(path.join(src, "index.ts"), "// code");
      await fs.writeFile(path.join(src, "src", "app.ts"), "// app");
      await fs.writeFile(path.join(src, "node_modules", "dep", "index.js"), "// dep");
      await fs.writeFile(path.join(src, "env", ".env.dev"), "FOO=bar");
      await fs.writeFile(path.join(src, "README.md"), "# readme");
      await fs.writeFile(path.join(src, "package-lock.json"), "{}");

      await copyExistMetaOSProject(src, dest);

      // Copied
      const indexContent = await fs.readFile(path.join(dest, "index.ts"), "utf-8");
      expect(indexContent).to.equal("// code");
      const appContent = await fs.readFile(path.join(dest, "src", "app.ts"), "utf-8");
      expect(appContent).to.equal("// app");

      // Excluded
      for (const excluded of ["node_modules", "env", "README.md", "package-lock.json"]) {
        try {
          await fs.access(path.join(dest, excluded));
          expect.fail(`${excluded} should have been excluded`);
        } catch {
          // expected - file/dir should not exist
        }
      }
    });
  });

  describe("unifyProjectID", () => {
    it("should update manifest.json id and .env.dev TEAMS_APP_ID", async () => {
      const project = path.join(tmpDir, "project");
      await fs.mkdir(path.join(project, "appPackage"), { recursive: true });
      await fs.mkdir(path.join(project, "env"), { recursive: true });

      await fs.writeFile(
        path.join(project, "appPackage", "manifest.json"),
        JSON.stringify({ id: "old-id", name: { short: "test" } })
      );
      await fs.writeFile(
        path.join(project, "env", ".env.dev"),
        "TEAMSFX_ENV=dev\nTEAMS_APP_ID=old-id\n"
      );

      await unifyProjectID(project);

      const manifest = JSON.parse(
        await fs.readFile(path.join(project, "appPackage", "manifest.json"), "utf-8")
      );
      expect(manifest.id).to.not.equal("old-id");
      expect(manifest.id).to.match(
        /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/
      );

      const envContent = await fs.readFile(path.join(project, "env", ".env.dev"), "utf-8");
      expect(envContent).to.include(`TEAMS_APP_ID=${manifest.id}`);
      expect(envContent).to.not.include("old-id");
    });

    it("should create env file if it does not exist", async () => {
      const project = path.join(tmpDir, "project2");
      await fs.mkdir(path.join(project, "appPackage"), { recursive: true });
      await fs.writeFile(
        path.join(project, "appPackage", "manifest.json"),
        JSON.stringify({ id: "old" })
      );

      await unifyProjectID(project);

      const envContent = await fs.readFile(path.join(project, "env", ".env.dev"), "utf-8");
      expect(envContent).to.include("TEAMS_APP_ID=");
    });
  });

  describe("extendToDA", () => {
    it("should create DA and action files and modify manifest", async () => {
      const project = path.join(tmpDir, "project3");
      await fs.mkdir(path.join(project, "appPackage"), { recursive: true });
      await fs.mkdir(path.join(project, "src", "commands"), { recursive: true });

      // Create a minimal manifest with extensions/runtimes
      const manifest = {
        id: "test-id",
        extensions: [
          {
            runtimes: [
              {
                code: { script: "commands.js" },
                actions: [],
              },
            ],
          },
        ],
      };
      await fs.writeFile(
        path.join(project, "appPackage", "manifest.json"),
        JSON.stringify(manifest)
      );
      await fs.writeFile(
        path.join(project, "src", "commands", "commands.ts"),
        "// existing code\n"
      );
      await fs.writeFile(
        path.join(project, "package.json"),
        JSON.stringify({ devDependencies: {} })
      );

      await extendToDA(project, "TestApp");

      // Check manifest was updated
      const updatedManifest = JSON.parse(
        await fs.readFile(path.join(project, "appPackage", "manifest.json"), "utf-8")
      );
      expect(updatedManifest.copilotAgents).to.exist;
      expect(updatedManifest.copilotAgents.declarativeAgents).to.be.an("array").with.length(1);

      // Check DA file was created
      const daFile = await fs.readFile(
        path.join(project, "appPackage", "declarativeAgent.json"),
        "utf-8"
      );
      const daJson = JSON.parse(daFile);
      expect(daJson.name).to.include("TestApp");

      // Check action file was created
      const actionFile = await fs.readFile(
        path.join(project, "appPackage", "alchemy-plugin.json"),
        "utf-8"
      );
      const actionJson = JSON.parse(actionFile);
      expect(actionJson.namespace).to.equal("AddInFunctions");

      // Check commands were appended
      const cmdsContent = await fs.readFile(
        path.join(project, "src", "commands", "commands.ts"),
        "utf-8"
      );
      expect(cmdsContent).to.include("addFooter");
      expect(cmdsContent).to.include("fillColor");
      expect(cmdsContent).to.include("addTextToSlide");

      // Check package.json was updated
      const pkg = JSON.parse(await fs.readFile(path.join(project, "package.json"), "utf-8"));
      expect(pkg.devDependencies["office-addin-debugging"]).to.equal("6.0.6");
    });

    it("should throw if manifest has no runtimes", async () => {
      const project = path.join(tmpDir, "project4");
      await fs.mkdir(path.join(project, "appPackage"), { recursive: true });
      await fs.mkdir(path.join(project, "src", "commands"), { recursive: true });

      const manifest = { id: "test-id", extensions: [{}] };
      await fs.writeFile(
        path.join(project, "appPackage", "manifest.json"),
        JSON.stringify(manifest)
      );
      await fs.writeFile(path.join(project, "src", "commands", "commands.ts"), "");
      await fs.writeFile(
        path.join(project, "package.json"),
        JSON.stringify({ devDependencies: {} })
      );

      try {
        await extendToDA(project, "TestApp");
        expect.fail("Should have thrown");
      } catch (e: any) {
        expect(e.message).to.include("No runtimes found");
      }
    });
  });
});
