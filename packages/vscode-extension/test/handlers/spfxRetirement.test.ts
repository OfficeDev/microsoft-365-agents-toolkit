// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import * as fs from "fs-extra";
import path from "path";
import { expect, it } from "vitest";
import packageJson from "../../package.json";

it("RET-SPFX-01: removes the obsolete contributed and registered add-web-part command", async () => {
  expect(packageJson.contributes.commands.map((command) => command.command)).not.to.contain(
    "fx-extension.addWebpart"
  );
  const registration = await fs.readFile(
    path.join(__dirname, "..", "..", "src", "extension.ts"),
    "utf8"
  );
  expect(registration).not.to.contain('"fx-extension.addWebpart"');
});
