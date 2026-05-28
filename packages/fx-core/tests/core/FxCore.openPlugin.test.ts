// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { err, ok, UserError } from "@microsoft/teamsfx-api";
import { assert } from "chai";
import "mocha";
import sinon from "sinon";
import { featureFlagManager, FeatureFlags } from "../../src/common/featureFlags";
import * as exporter from "../../src/component/generator/openPlugin/exporter";
import * as importer from "../../src/component/generator/openPlugin/importer";
import { FxCore } from "../../src/core/FxCore";
import { setTools } from "../../src/common/globalVars";
import { MockTools } from "./utils";

describe("FxCore.openPlugin", () => {
  const sandbox = sinon.createSandbox();
  let core: FxCore;

  beforeEach(() => {
    setTools(new MockTools());
    core = new FxCore(new MockTools());
    sandbox
      .stub(featureFlagManager, "getBooleanValue")
      .callsFake((flag) => flag.name === FeatureFlags.OpenPluginImportExport.name);
  });

  afterEach(() => {
    sandbox.restore();
  });

  describe("importOpenPlugin", () => {
    it("returns UserError when feature flag is disabled", async () => {
      sandbox.restore();
      sandbox.stub(featureFlagManager, "getBooleanValue").returns(false);
      const res = await core.importOpenPlugin({ platform: "cli", path: "x" } as any);
      assert.isTrue(res.isErr());
      if (res.isErr()) assert.equal(res.error.name, "FeatureFlagDisabled");
    });

    it("returns MissingRequiredInput when --path is absent", async () => {
      const res = await core.importOpenPlugin({ platform: "cli" } as any);
      assert.isTrue(res.isErr());
      if (res.isErr()) assert.equal(res.error.name, "MissingRequiredInput");
    });

    it("returns InvalidDefaultAuthType for an unknown auth type", async () => {
      const res = await core.importOpenPlugin({
        platform: "cli",
        path: "/tmp",
        "default-auth-type": "Bogus",
      } as any);
      assert.isTrue(res.isErr());
      if (res.isErr()) assert.equal(res.error.name, "InvalidDefaultAuthType");
    });

    it("delegates to importer.importOpenPlugin on the success path", async () => {
      const stub = sandbox
        .stub(importer, "importOpenPlugin")
        .resolves(ok({ projectPath: "/tmp/out", warnings: ["w"] }));
      const res = await core.importOpenPlugin({
        platform: "cli",
        path: "/tmp/in",
        output: "/tmp/out",
        "privacy-url": "https://x/p",
        "terms-url": "https://x/t",
      } as any);
      assert.isTrue(res.isOk());
      if (res.isOk()) {
        assert.equal(res.value.projectPath, "/tmp/out");
        assert.equal(res.value.warnings?.[0].type, "openPluginImport");
        assert.equal(res.value.warnings?.[0].content, "w");
      }
      assert.isTrue(stub.calledOnce);
      const arg = stub.firstCall.args[0];
      assert.equal(arg.path, "/tmp/in");
      assert.equal(arg.privacyUrl, "https://x/p");
      assert.equal(arg.termsUrl, "https://x/t");
    });

    it("propagates importer errors", async () => {
      sandbox
        .stub(importer, "importOpenPlugin")
        .resolves(err(new UserError("OpenPluginImport", "Boom", "boom")));
      const res = await core.importOpenPlugin({
        platform: "cli",
        path: "/tmp/in",
      } as any);
      assert.isTrue(res.isErr());
    });
  });

  describe("exportOpenPlugin", () => {
    it("returns UserError when feature flag is disabled", async () => {
      sandbox.restore();
      sandbox.stub(featureFlagManager, "getBooleanValue").returns(false);
      const res = await core.exportOpenPlugin({ platform: "cli", path: "x" } as any);
      assert.isTrue(res.isErr());
      if (res.isErr()) assert.equal(res.error.name, "FeatureFlagDisabled");
    });

    it("returns MissingRequiredInput when --path is absent", async () => {
      const res = await core.exportOpenPlugin({ platform: "cli" } as any);
      assert.isTrue(res.isErr());
      if (res.isErr()) assert.equal(res.error.name, "MissingRequiredInput");
    });

    it("returns InvalidManifestKind for an unknown kind", async () => {
      const res = await core.exportOpenPlugin({
        platform: "cli",
        path: "/tmp",
        "manifest-kind": "bogus",
      } as any);
      assert.isTrue(res.isErr());
      if (res.isErr()) assert.equal(res.error.name, "InvalidManifestKind");
    });

    it("delegates to exporter.exportOpenPlugin on the success path", async () => {
      const stub = sandbox
        .stub(exporter, "exportOpenPlugin")
        .resolves(ok({ outputPath: "/tmp/out", warnings: ["w"] }));
      const res = await core.exportOpenPlugin({
        platform: "cli",
        path: "/tmp/proj",
        output: "/tmp/out",
        "manifest-kind": "claude-plugin",
      } as any);
      assert.isTrue(res.isOk());
      if (res.isOk()) {
        assert.equal(res.value.outputPath, "/tmp/out");
        assert.equal(res.value.warnings[0].type, "openPluginExport");
      }
      assert.isTrue(stub.calledOnce);
      const arg = stub.firstCall.args[0];
      assert.equal(arg.path, "/tmp/proj");
      assert.equal(arg.manifestKind, "claude-plugin");
    });

    it("propagates exporter errors", async () => {
      sandbox
        .stub(exporter, "exportOpenPlugin")
        .resolves(err(new UserError("OpenPluginExport", "Boom", "boom")));
      const res = await core.exportOpenPlugin({
        platform: "cli",
        path: "/tmp/proj",
      } as any);
      assert.isTrue(res.isErr());
    });
  });
});
