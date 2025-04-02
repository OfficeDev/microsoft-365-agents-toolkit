// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import { engine } from "./engine";
import { rootCommand } from "./models/root";

export async function start(binName: "teamsfx" | "teamsapp" = "teamsapp"): Promise<void> {
  rootCommand.name = binName;
  rootCommand.fullName = binName;
  await engine.start(rootCommand);
  // process.exit(0);
}
