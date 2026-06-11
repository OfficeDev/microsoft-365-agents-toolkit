// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { FxCore } from "@microsoft/teamsfx-core";
import { assert } from "chai";
import { envAddDeps } from "../../src/commands/models/envAdd";
import { envListDeps } from "../../src/commands/models/envList";

describe("CLI command deps wrappers", () => {
  it("envAddDeps delegates to projectSettingsHelper and activate", () => {
    const isValid = envAddDeps.isValidProjectV3(undefined);
    assert.isBoolean(isValid);

    const core = envAddDeps.getFxCore();
    assert.isTrue(core instanceof FxCore);
  });

  it("envListDeps delegates to projectSettingsHelper", () => {
    const isValid = envListDeps.isValidProjectV3(undefined);
    assert.isBoolean(isValid);
  });
});
