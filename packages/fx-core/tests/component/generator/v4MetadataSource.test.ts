// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { assert } from "chai";
import "mocha";
import { ok, Result } from "neverthrow";
import * as sinon from "sinon";
import { FxError } from "@microsoft/teamsfx-api";
import templateConfig from "../../../src/common/templates-config.json";
import { resolveV4MetadataSource } from "../../../src/component/generator/v4MetadataSource";
import * as bundledFloor from "../../../src/v4/distribution/bundledFloor";
import * as templateSource from "../../../src/v4/distribution/templateSource";
import * as templateSourcePort from "../../../src/v4/distribution/templateSourcePort";

describe("resolveV4MetadataSource", () => {
  const sandbox = sinon.createSandbox();

  afterEach(() => {
    sandbox.restore();
  });

  it("resolves through the same single decision point as the template package", async () => {
    const fakeFloor = { version: "6.10.1" } as any;
    const fakePort = { tagList: () => Promise.resolve([]) } as any;
    const loadFloorStub = sandbox.stub(bundledFloor, "loadBundledFloor").returns(fakeFloor);
    const createPortStub = sandbox
      .stub(templateSourcePort, "createTemplateSourcePort")
      .returns(fakePort);
    const expected: Result<templateSource.TemplateSource, FxError> = ok({
      origin: "online" as const,
      version: "2.0.0",
      digest: "sha256:x",
      location: "",
    });
    const resolveStub = sandbox.stub(templateSource, "resolveTemplateSource").resolves(expected);

    const result = await resolveV4MetadataSource();

    assert.strictEqual(result, expected);
    assert.isTrue(loadFloorStub.calledOnce);
    // The port is built from the v4 channel config in templates-config.json.
    assert.isTrue(
      createPortStub.calledOnceWith(
        sinon.match({
          templatesV4TagListURL: templateConfig.templatesV4TagListURL,
          templateDownloadBaseURL: templateConfig.templateDownloadBaseURL,
        }),
        fakeFloor
      )
    );
    // The resolver reads v4.range / v4.bundled (never the top-level version).
    assert.isTrue(
      resolveStub.calledOnceWith(
        sinon.match({
          range: templateConfig.v4.range,
          bundled: templateConfig.v4.bundled,
          port: fakePort,
        })
      )
    );
  });
});
