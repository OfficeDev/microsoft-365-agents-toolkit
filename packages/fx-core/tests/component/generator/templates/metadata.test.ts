// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { Platform } from "@microsoft/teamsfx-api";
import { assert } from "chai";
import fs from "fs-extra";
import * as path from "path";
import * as sinon from "sinon";
import { featureFlagManager } from "../../../../src/common/featureFlags";
import * as templateHelper from "../../../../src/component/generator/templateHelper";
import {
  getAllTemplatesOnPlatform,
  getDefaultTemplatesOnPlatform,
} from "../../../../src/component/generator/templates/metadata";
import { Template } from "../../../../src/component/generator/templates/metadata/interface";
import * as folder from "../../../../src/folder";

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
      sandbox.stub(folder, "getTemplatesFolder").returns(path.resolve("/bundled"));
      sandbox.stub(fs, "pathExistsSync").returns(false);
      const readFileSyncStub = sandbox
        .stub(fs, "readFileSync")
        .returns(JSON.stringify(mockTemplates));

      getAllTemplatesOnPlatform(Platform.VS);

      const readPath = readFileSyncStub.firstCall.args[0] as string;
      assert.include(readPath, path.join("metadata", "allTemplates.json"));
    });

    it("falls back to bundled path when useLocalTemplate is true", () => {
      sandbox.stub(templateHelper, "useLocalTemplate").returns(true);
      sandbox.stub(folder, "getTemplatesFolder").returns(path.resolve("/bundled"));
      sandbox.stub(fs, "pathExistsSync").returns(true);
      const readFileSyncStub = sandbox
        .stub(fs, "readFileSync")
        .returns(JSON.stringify(mockTemplates));

      getAllTemplatesOnPlatform(Platform.VS);

      const readPath = readFileSyncStub.firstCall.args[0] as string;
      assert.include(readPath, path.join("metadata", "allTemplates.json"));
    });

    it("falls back to bundled path when v4 channel forces bundled metadata even if cache exists", () => {
      sandbox.stub(templateHelper, "useLocalTemplate").returns(false);
      sandbox.stub(featureFlagManager, "getBooleanValue").returns(true);
      sandbox.stub(fs, "pathExistsSync").callsFake((p: fs.PathLike) => {
        const value = String(p);
        // Simulate v4 channel with no downloaded v4 cache marker.
        if (value.endsWith("template-version-v4.txt")) {
          return false;
        }
        return true;
      });
      const readFileSyncStub = sandbox
        .stub(fs, "readFileSync")
        .returns(JSON.stringify(mockTemplates));

      getAllTemplatesOnPlatform(Platform.VSCode);

      const readPath = readFileSyncStub.firstCall.args[0] as string;
      assert.notInclude(readPath, ".fx");
      assert.include(readPath, path.join("metadata", "allTemplates.json"));
    });

    it("keeps reading the VS cache even when v4 channel forces bundled metadata", () => {
      sandbox.stub(templateHelper, "useLocalTemplate").returns(false);
      sandbox.stub(templateHelper, "useBundledMetadataForV4").returns(true);
      sandbox.stub(folder, "getTemplatesFolder").returns("/bundled");
      const pathExistsStub = sandbox.stub(fs, "pathExistsSync").returns(true);
      sandbox.stub(fs, "readFileSync").returns(JSON.stringify(mockTemplates));

      getAllTemplatesOnPlatform(Platform.VS);

      // The v4 migration covers only VSC/CLI; VS keeps its v3 vs-metadata cache.
      const checkedPath = pathExistsStub.firstCall.args[0] as string;
      assert.include(checkedPath, "vs-metadata");
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

describe("useBundledMetadataForV4", () => {
  const sandbox = sinon.createSandbox();

  afterEach(() => {
    sandbox.restore();
  });

  it("returns false when the v4 flag is off", () => {
    sandbox.stub(featureFlagManager, "getBooleanValue").returns(false);
    const pathExistsStub = sandbox.stub(fs, "pathExistsSync").returns(false);

    assert.isFalse(templateHelper.useBundledMetadataForV4());
    // Short-circuits before touching the filesystem.
    assert.isFalse(pathExistsStub.called);
  });

  it("returns false (read the downloaded v4 cache) when the v4 version file exists", () => {
    sandbox.stub(featureFlagManager, "getBooleanValue").returns(true);
    const pathExistsStub = sandbox.stub(fs, "pathExistsSync").returns(true);

    assert.isFalse(templateHelper.useBundledMetadataForV4());
    const checkedPath = pathExistsStub.firstCall.args[0] as string;
    assert.include(checkedPath, "template-version-v4.txt");
  });

  it("returns true (read bundled) when the v4 version file is absent", () => {
    sandbox.stub(featureFlagManager, "getBooleanValue").returns(true);
    sandbox.stub(fs, "pathExistsSync").returns(false);

    assert.isTrue(templateHelper.useBundledMetadataForV4());
  });
});
