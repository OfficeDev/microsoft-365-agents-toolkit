// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { err, FxError, ok, Result, UserError, Void } from "@microsoft/teamsfx-api";
import { getLocalizedString } from "../common/localizeUtils";
import { SolutionSource } from "./constants";
import { DriverContext } from "./driver/interface/commonArgs";

class DeployUtils {
  async askForDeployConsentV3(ctx: DriverContext): Promise<Result<Void, FxError>> {
    const msg = getLocalizedString("core.deploy.confirmEnvNoticeV3", process.env.TEAMSFX_ENV);
    const deployOption = getLocalizedString("core.option.deploy");
    const result = await ctx.ui?.showMessage("warn", msg, true, deployOption);
    const choice = result?.isOk() ? result.value : undefined;
    if (choice === deployOption) {
      return ok(Void);
    }
    return err(new UserError(SolutionSource, "UserCancel", "UserCancel"));
  }
}
export const deployUtils = new DeployUtils();
