// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import { logger } from "../commonlib/logger";
import { engine } from "./engine";
import { rootCommand } from "./models/root";

export async function start(binName: "teamsfx" | "teamsapp" = "teamsapp"): Promise<void> {
  rootCommand.name = binName;
  rootCommand.fullName = binName;

  logger.warning(
    `Deprecation Warning: The CLI package "@microsoft/teamsapp-cli" is being renamed to "@microsoft/m365agentstoolkit-cli". The command name "teamsapp" will also change to "m365agents" accordingly in the upcoming release. The existing package will no longer receive updates. To stay up to date, please switch to the new package and update your workflows accordingly once it is available. `
  );

  await engine.start(rootCommand);
  // process.exit(0);
}
