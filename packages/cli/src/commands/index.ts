// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import { logger } from "../commonlib/logger";
import { engine } from "./engine";
import { rootCommand } from "./models/root";

export async function start(binName: "teamsfx" | "teamsapp" = "teamsapp"): Promise<void> {
  rootCommand.name = binName;
  rootCommand.fullName = binName;
  logger.warning(
    `Deprecation Warning: The CLI package "@microsoft/teamsapp-cli" and its "teamsapp" command are being renamed in the upcoming release. The existing package will no longer receive updates. Please switch to the new package and update your workflows accordingly once it's available.`
  );
  await engine.start(rootCommand);
  // process.exit(0);
}
