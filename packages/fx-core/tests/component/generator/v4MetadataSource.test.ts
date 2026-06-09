// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { FxError } from "@microsoft/teamsfx-api";
import { assert } from "chai";
import { Result } from "neverthrow";
import templateConfig from "../../../src/common/templates-config.json";
import { resolveV4MetadataSource } from "../../../src/component/generator/v4MetadataSource";
import * as bundledFloor from "../../../src/v4/distribution/bundledFloor";
import * as templateSource from "../../../src/v4/distribution/templateSource";
import * as templateSourcePort from "../../../src/v4/distribution/templateSourcePort";

describe("resolveV4MetadataSource", () => {
  it("resolves through the same single decision point as the template package", async () => {
    const floor = bundledFloor.loadBundledFloor();
    const port = templateSourcePort.createTemplateSourcePort(
      {
        templatesV4TagListURL: templateConfig.templatesV4TagListURL,
        templateDownloadBaseURL: templateConfig.templateDownloadBaseURL,
      },
      floor
    );
    const expected: Result<templateSource.TemplateSource, FxError> =
      await templateSource.resolveTemplateSource({
        range: templateConfig.v4.range,
        bundled: templateConfig.v4.bundled,
        port,
      });

    const result = await resolveV4MetadataSource();

    assert.deepEqual(result, expected);
  });
});
