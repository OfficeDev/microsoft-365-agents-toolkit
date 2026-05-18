// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { hooks } from "@feathersjs/hooks/lib";
import {
  CreateProjectResult,
  err,
  FxError,
  Inputs,
  ok,
  Result,
  Stage,
  UserError,
  Warning,
} from "@microsoft/teamsfx-api";
import { featureFlagManager, FeatureFlags } from "../common/featureFlags";
import { ErrorContextMW } from "../common/globalVars";
import {
  convertOpenPlugin,
  OPEN_PLUGIN_CONVERT_SOURCE,
} from "../component/generator/openPlugin/generator";
import { ConvertInputs, DefaultAuthOption } from "../component/generator/openPlugin/types";
import { ErrorHandlerMW } from "./middleware/errorHandler";
import { FxCoreDeclarativeAgentPart } from "./FxCore.declarativeAgent";

const ALLOWED_AUTH: DefaultAuthOption[] = [
  "Auto",
  "None",
  "OAuthPluginVault",
  "ApiKeyPluginVault",
];

export class FxCoreOpenPluginPart extends FxCoreDeclarativeAgentPart {
  /**
   * Convert an Open Plugin (Open Plugin Spec v1.0 / Claude Code plugin /
   * Cursor plugin) directory into a scaffolded Teams Toolkit project.
   */
  @hooks([
    ErrorContextMW({ component: "FxCore", stage: Stage.create, reset: true }),
    ErrorHandlerMW,
  ])
  async convertOpenPlugin(inputs: Inputs): Promise<Result<CreateProjectResult, FxError>> {
    if (!featureFlagManager.getBooleanValue(FeatureFlags.OpenPluginConvert)) {
      return err(
        new UserError(
          OPEN_PLUGIN_CONVERT_SOURCE,
          "FeatureFlagDisabled",
          `Set ${FeatureFlags.OpenPluginConvert.name}=true to enable 'atk convert openplugin'.`
        )
      );
    }

    const validatedInputs = this.validateConvertInputs(inputs);
    if (validatedInputs.isErr()) {
      return err(validatedInputs.error);
    }
    const res = await convertOpenPlugin(validatedInputs.value);
    if (res.isErr()) {
      return err(res.error);
    }
    const warnings: Warning[] = res.value.warnings.map((content) => ({
      type: "openPluginConvert",
      content,
    }));
    return ok({ projectPath: res.value.projectPath, warnings });
  }

  private validateConvertInputs(inputs: Inputs): Result<ConvertInputs, FxError> {
    const pluginPath = inputs["path"];
    const privacyUrl = inputs["privacy-url"];
    const termsUrl = inputs["terms-url"];

    const missing: string[] = [];
    if (!pluginPath) missing.push("path");
    if (!privacyUrl) missing.push("privacy-url");
    if (!termsUrl) missing.push("terms-url");
    if (missing.length > 0) {
      return err(
        new UserError(
          OPEN_PLUGIN_CONVERT_SOURCE,
          "MissingRequiredInput",
          `Missing required option(s): ${missing.join(", ")}.`
        )
      );
    }

    let defaultAuthType: DefaultAuthOption | undefined;
    const rawAuth = inputs["default-auth-type"];
    if (rawAuth) {
      if (!ALLOWED_AUTH.includes(rawAuth as DefaultAuthOption)) {
        return err(
          new UserError(
            OPEN_PLUGIN_CONVERT_SOURCE,
            "InvalidDefaultAuthType",
            `--default-auth-type must be one of: ${ALLOWED_AUTH.join(", ")}.`
          )
        );
      }
      defaultAuthType = rawAuth as DefaultAuthOption;
    }

    return ok({
      path: pluginPath as string,
      output: inputs["output"] as string | undefined,
      privacyUrl: privacyUrl as string,
      termsUrl: termsUrl as string,
      websiteUrl: inputs["website-url"] as string | undefined,
      appId: inputs["app-id"] as string | undefined,
      defaultAuthType,
      packageName: inputs["package-name"] as string | undefined,
    });
  }
}
