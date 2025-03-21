// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import { CLICommand, FxError, Result, err, ok, signedOut } from "@microsoft/teamsfx-api";
import {
  AppStudioScopes,
  CheckerFactory,
  DepsType,
  EmptyLogger,
  EmptyTelemetry,
  FuncToolChecker,
  LocalCertificateManager,
  assembleError,
  getSideloadingStatus,
} from "@microsoft/teamsfx-core";
import * as util from "util";
import { getFxCore } from "../../../activate";
import { DoneText, TextType, WarningText, colorize } from "../../../colorize";
import { logger } from "../../../commonlib/logger";
import M365TokenInstance from "../../../commonlib/m365Login";
import { cliSource } from "../../../constants";
import { commands, strings } from "../../../resource";
import { TelemetryEvent } from "../../../telemetry/cliTelemetryEvents";

export const teamsappDoctorCommand: CLICommand = {
  name: "doctor",
  description: commands.doctor.description,
  options: [],
  telemetry: {
    event: TelemetryEvent.Doctor,
  },
  defaultInteractiveOption: false,
  handler: async () => {
    getFxCore();
    const checker = new DoctorChecker();
    await checker.checkAccount();
    await checker.checkNodejs();
    await checker.checkFuncCoreTool();
    await checker.checkCert();
    return ok(undefined);
  },
};

export class DoctorChecker {
  async checkAccount(): Promise<void> {
    const res = await this.checkM365Account();
    if (res.isErr()) {
      logger.error(res.error.message);
    } else {
      logger.info(res.value);
    }
  }

  async checkNodejs(): Promise<void> {
    const nodeChecker = CheckerFactory.createChecker(
      DepsType.LtsNode,
      new EmptyLogger(),
      new EmptyTelemetry()
    );
    const nodeRes = await nodeChecker.getInstallationInfo();
    if (nodeRes.isInstalled) {
      if (nodeRes.error) {
        logger.info(
          WarningText +
            util.format(
              strings.command.doctor.node.NotSupported,
              nodeRes.details?.installVersion,
              nodeRes.details?.supportedVersions.join(", ")
            )
        );
      } else {
        logger.info(
          DoneText +
            util.format(strings.command.doctor.node.Success, nodeRes.details?.installVersion)
        );
      }
    } else {
      logger.info(WarningText + strings.command.doctor.node.NotFound);
    }
  }

  async checkFuncCoreTool(): Promise<void> {
    const funcChecker = new FuncToolChecker();
    try {
      const funcRes = await funcChecker.queryFuncVersion(undefined);
      logger.info(DoneText + util.format(strings.command.doctor.func.Success, funcRes.versionStr));
    } catch (e) {
      logger.info(WarningText + strings.command.doctor.func.NotFound);
    }
  }

  async checkCert(): Promise<void> {
    const certManager = new LocalCertificateManager();
    const certRes = await certManager.setupCertificate(true, true);
    if (!certRes.found) {
      logger.info(WarningText + strings.command.doctor.cert.NotFound);
    } else {
      if (certRes.alreadyTrusted) {
        logger.info(DoneText + strings.command.doctor.cert.Success);
      } else {
        logger.info(WarningText + strings.command.doctor.cert.FoundNotTrust);
      }
    }
  }

  async checkM365Account(): Promise<Result<string, FxError>> {
    let result = true;
    let summaryMsg = "";
    let error = undefined;
    let loginHint: string | undefined = undefined;
    try {
      let loginStatusRes = await M365TokenInstance.getStatus({ scopes: AppStudioScopes });
      let token = loginStatusRes.isOk() ? loginStatusRes.value.token : undefined;
      if (loginStatusRes.isOk() && loginStatusRes.value.status === signedOut) {
        const tokenRes = await M365TokenInstance.getAccessToken({
          scopes: AppStudioScopes,
          showDialog: true,
        });
        token = tokenRes.isOk() ? tokenRes.value : undefined;
        loginStatusRes = await M365TokenInstance.getStatus({ scopes: AppStudioScopes });
      }
      if (token === undefined) {
        result = false;
        summaryMsg = WarningText + strings.command.doctor.account.NotSignIn;
      } else {
        const isSideloadingEnabled = await getSideloadingStatus(token);
        if (isSideloadingEnabled === false) {
          // sideloading disabled
          result = false;
          summaryMsg = WarningText + strings.command.doctor.account.SideLoadingDisabled;
        }
      }
      const tokenObject = loginStatusRes.isOk() ? loginStatusRes.value.accountInfo : undefined;
      if (tokenObject && tokenObject.upn) {
        loginHint = tokenObject.upn as string;
      }
    } catch (e) {
      error = assembleError(e, cliSource);
      return err(error);
    }
    if (result && loginHint) {
      summaryMsg =
        DoneText +
        util.format(
          strings.command.doctor.account.SignInSuccess,
          colorize(loginHint, TextType.Email)
        );
    }
    return ok(summaryMsg);
  }
}
