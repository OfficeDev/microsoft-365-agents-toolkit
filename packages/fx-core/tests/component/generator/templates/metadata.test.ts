// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { Platform } from "@microsoft/teamsfx-api";
import { assert } from "chai";
import fs from "fs-extra";
import "mocha";
import * as path from "path";
import * as sinon from "sinon";
import {
  getAllTemplatesOnPlatform,
  getDefaultTemplatesOnPlatform,
} from "../../../../src/component/generator/templates/metadata";
import * as templateHelper from "../../../../src/component/generator/templateHelper";
import * as folder from "../../../../src/folder";
import { Template } from "../../../../src/component/generator/templates/metadata/interface";

const mockTemplates: Template[] = [
  { id: "t1", name: "TypeScript Bot", language: "typescript", description: "A TS bot" },
  { id: "t2", name: "CSharp Bot", language: "csharp", description: "A C# bot" },
  { id: "t3", name: "JavaScript Tab", language: "javascript", description: "A JS tab" },
];

describe("metadata platform routing", () => {
  const sandbox = sinon.createSandbox();

  afterEach(() => {
    sandbox.restore();
  });

  describe("getAllTemplatesOnPlatform", () => {
    it("reads from vs-metadata subdir when cache exists for Platform.VS", () => {
      sandbox.stub(templateHelper, "useLocalTemplate").returns(false);
      sandbox.stub(folder, "getTemplatesFolder").returns("/bundled");
      const pathExistsStub = sandbox.stub(fs, "pathExistsSync").returns(true);
      sandbox.stub(fs, "readFileSync").returns(JSON.stringify(mockTemplates));

      getAllTemplatesOnPlatform(Platform.VS);

      const checkedPath = pathExistsStub.firstCall.args[0] as string;
      assert.include(checkedPath, "vs-metadata");
      assert.include(checkedPath, "allTemplates.json");
    });

    it("reads from metadata subdir when cache exists for Platform.VSCode", () => {
      sandbox.stub(templateHelper, "useLocalTemplate").returns(false);
      sandbox.stub(folder, "getTemplatesFolder").returns("/bundled");
      const pathExistsStub = sandbox.stub(fs, "pathExistsSync").returns(true);
      sandbox.stub(fs, "readFileSync").returns(JSON.stringify(mockTemplates));

      getAllTemplatesOnPlatform(Platform.VSCode);

      const checkedPath = pathExistsStub.firstCall.args[0] as string;
      assert.notInclude(checkedPath, "vs-metadata");
      assert.include(checkedPath, path.join(".fx", "metadata"));
      assert.include(checkedPath, "allTemplates.json");
    });

    it("reads from metadata subdir when cache exists for Platform.CLI", () => {
      sandbox.stub(templateHelper, "useLocalTemplate").returns(false);
      sandbox.stub(folder, "getTemplatesFolder").returns("/bundled");
      const pathExistsStub = sandbox.stub(fs, "pathExistsSync").returns(true);
      sandbox.stub(fs, "readFileSync").returns(JSON.stringify(mockTemplates));

      getAllTemplatesOnPlatform(Platform.CLI);

      const checkedPath = pathExistsStub.firstCall.args[0] as string;
      assert.notInclude(checkedPath, "vs-metadata");
    });

    it("falls back to bundled path when cache does not exist", () => {
      sandbox.stub(templateHelper, "useLocalTemplate").returns(false);
      const bundledPath = path.resolve("/bundled");
      sandbox.stub(folder, "getTemplatesFolder").returns(bundledPath);
      sandbox.stub(fs, "pathExistsSync").returns(false);
      const readFileSyncStub = sandbox
        .stub(fs, "readFileSync")
        .returns(JSON.stringify(mockTemplates));

      getAllTemplatesOnPlatform(Platform.VS);

      const readPath = readFileSyncStub.firstCall.args[0] as string;
      assert.include(readPath, bundledPath);
      assert.notInclude(readPath, "vs-metadata");
    });

    it("falls back to bundled path when useLocalTemplate is true", () => {
      sandbox.stub(templateHelper, "useLocalTemplate").returns(true);
      const bundledPath = path.resolve("/bundled");
      sandbox.stub(folder, "getTemplatesFolder").returns(bundledPath);
      sandbox.stub(fs, "pathExistsSync").returns(true);
      const readFileSyncStub = sandbox
        .stub(fs, "readFileSync")
        .returns(JSON.stringify(mockTemplates));

      getAllTemplatesOnPlatform(Platform.VS);

      const readPath = readFileSyncStub.firstCall.args[0] as string;
      assert.include(readPath, bundledPath);
    });

    it("returns only csharp templates for Platform.VS", () => {
      sandbox.stub(templateHelper, "useLocalTemplate").returns(false);
      sandbox.stub(folder, "getTemplatesFolder").returns("/bundled");
      sandbox.stub(fs, "pathExistsSync").returns(true);
      sandbox.stub(fs, "readFileSync").returns(JSON.stringify(mockTemplates));

      const result = getAllTemplatesOnPlatform(Platform.VS);

      assert.deepEqual(result, [mockTemplates[1]]);
    });

    it("returns only non-csharp templates for Platform.VSCode", () => {
      sandbox.stub(templateHelper, "useLocalTemplate").returns(false);
      sandbox.stub(folder, "getTemplatesFolder").returns("/bundled");
      sandbox.stub(fs, "pathExistsSync").returns(true);
      sandbox.stub(fs, "readFileSync").returns(JSON.stringify(mockTemplates));

      const result = getAllTemplatesOnPlatform(Platform.VSCode);

      assert.deepEqual(result, [mockTemplates[0], mockTemplates[2]]);
    });

    it("returns all templates for Platform.CLI", () => {
      sandbox.stub(templateHelper, "useLocalTemplate").returns(false);
      sandbox.stub(folder, "getTemplatesFolder").returns("/bundled");
      sandbox.stub(fs, "pathExistsSync").returns(true);
      sandbox.stub(fs, "readFileSync").returns(JSON.stringify(mockTemplates));

      const result = getAllTemplatesOnPlatform(Platform.CLI);

      assert.deepEqual(result, mockTemplates);
    });

    it("returns empty array for unknown platform", () => {
      sandbox.stub(templateHelper, "useLocalTemplate").returns(false);
      sandbox.stub(folder, "getTemplatesFolder").returns("/bundled");
      sandbox.stub(fs, "pathExistsSync").returns(true);
      sandbox.stub(fs, "readFileSync").returns(JSON.stringify(mockTemplates));

      const result = getAllTemplatesOnPlatform("unknown" as Platform);

      assert.deepEqual(result, []);
    });
  });

  describe("getDefaultTemplatesOnPlatform", () => {
    it("reads from vs-metadata subdir when cache exists for Platform.VS", () => {
      sandbox.stub(templateHelper, "useLocalTemplate").returns(false);
      sandbox.stub(folder, "getTemplatesFolder").returns("/bundled");
      const pathExistsStub = sandbox.stub(fs, "pathExistsSync").returns(true);
      sandbox.stub(fs, "readFileSync").returns(JSON.stringify(mockTemplates));

      getDefaultTemplatesOnPlatform(Platform.VS);

      const checkedPath = pathExistsStub.firstCall.args[0] as string;
      assert.include(checkedPath, "vs-metadata");
      assert.include(checkedPath, "defaultGeneratorTemplates.json");
    });

    it("reads from metadata subdir when cache exists for Platform.VSCode", () => {
      sandbox.stub(templateHelper, "useLocalTemplate").returns(false);
      sandbox.stub(folder, "getTemplatesFolder").returns("/bundled");
      const pathExistsStub = sandbox.stub(fs, "pathExistsSync").returns(true);
      sandbox.stub(fs, "readFileSync").returns(JSON.stringify(mockTemplates));

      getDefaultTemplatesOnPlatform(Platform.VSCode);

      const checkedPath = pathExistsStub.firstCall.args[0] as string;
      assert.notInclude(checkedPath, "vs-metadata");
      assert.include(checkedPath, "defaultGeneratorTemplates.json");
    });

    it("returns only csharp templates for Platform.VS", () => {
      sandbox.stub(templateHelper, "useLocalTemplate").returns(false);
      sandbox.stub(folder, "getTemplatesFolder").returns("/bundled");
      sandbox.stub(fs, "pathExistsSync").returns(true);
      sandbox.stub(fs, "readFileSync").returns(JSON.stringify(mockTemplates));

      const result = getDefaultTemplatesOnPlatform(Platform.VS);

      assert.deepEqual(result, [mockTemplates[1]]);
    });

    it("returns only non-csharp templates for Platform.VSCode", () => {
      sandbox.stub(templateHelper, "useLocalTemplate").returns(false);
      sandbox.stub(folder, "getTemplatesFolder").returns("/bundled");
      sandbox.stub(fs, "pathExistsSync").returns(true);
      sandbox.stub(fs, "readFileSync").returns(JSON.stringify(mockTemplates));

      const result = getDefaultTemplatesOnPlatform(Platform.VSCode);

      assert.deepEqual(result, [mockTemplates[0], mockTemplates[2]]);
    });

    it("returns all templates for Platform.CLI", () => {
      sandbox.stub(templateHelper, "useLocalTemplate").returns(false);
      sandbox.stub(folder, "getTemplatesFolder").returns("/bundled");
      sandbox.stub(fs, "pathExistsSync").returns(true);
      sandbox.stub(fs, "readFileSync").returns(JSON.stringify(mockTemplates));

      const result = getDefaultTemplatesOnPlatform(Platform.CLI);

      assert.deepEqual(result, mockTemplates);
    });

    it("returns empty array for unknown platform", () => {
      sandbox.stub(templateHelper, "useLocalTemplate").returns(false);
      sandbox.stub(folder, "getTemplatesFolder").returns("/bundled");
      sandbox.stub(fs, "pathExistsSync").returns(true);
      sandbox.stub(fs, "readFileSync").returns(JSON.stringify(mockTemplates));

      const result = getDefaultTemplatesOnPlatform("unknown" as Platform);

      assert.deepEqual(result, []);
    });
  });
});
