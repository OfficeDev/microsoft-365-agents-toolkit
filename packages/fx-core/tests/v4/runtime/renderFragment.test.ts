// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { assert } from "vitest";
import { renderFragment } from "../../../src/v4/runtime/renderFragment";

describe("domain asset fragments", () => {
  it("AC-30: preserves raw code characters, environment references and line endings", () => {
    assert.strictEqual(
      renderFragment(["{{value}}", "${workspaceFolder}", ""], { value: 'A & <B> "C"' }),
      'A & <B> "C"\n${workspaceFolder}\n'
    );
  });
});
