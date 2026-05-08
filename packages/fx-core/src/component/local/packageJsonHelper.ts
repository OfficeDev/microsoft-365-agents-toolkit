// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
"use strict";

import { LogProvider } from "@microsoft/teamsfx-api";
import * as fs from "fs-extra";
import * as path from "path";

export async function loadPackageJson(path: string, logger?: LogProvider): Promise<any> {
  if (!(await fs.pathExists(path))) {
    logger?.error(`Cannot load package.json from ${path}. File not found.`);
    return undefined;
  }

  const rpj = require("read-package-json-fast");
  try {
    return await rpj(path);
  } catch (error) {
    logger?.error(`Cannot load package.json from ${path}. Error: ${error}`);
    return undefined;
  }
}

export async function loadTeamsFxDevScript(componentRoot: string): Promise<string | undefined> {
  const packageJson = await loadPackageJson(path.join(componentRoot, "package.json"));
  if (packageJson && packageJson.scripts && packageJson.scripts["dev:teamsfx"]) {
    const devTeamsfx: string = packageJson.scripts["dev:teamsfx"];
    return devTeamsfx;
  } else {
    return undefined;
  }
}
