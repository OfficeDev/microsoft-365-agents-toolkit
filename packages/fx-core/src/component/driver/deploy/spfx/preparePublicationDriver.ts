// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { hooks } from "@feathersjs/hooks/lib";
import {
  FxError,
  SystemError,
  SystemErrorOptions,
  UserError,
  UserErrorOptions,
  err,
  ok,
} from "@microsoft/teamsfx-api";
import { Service } from "typedi";
import { getLocalizedString } from "../../../../common/localizeUtils";
import { maskSecret } from "../../../../common/stringUtils";
import { ErrorCategory } from "../../../../error/types";
import { GraphScopes, SPFxScopes } from "../../../../common/constants";
import { getSPFxToken } from "../../../../common/tools";
import { DriverContext } from "../../interface/commonArgs";
import { ExecutionResult, StepDriver } from "../../interface/stepDriver";
import { addStartAndEndTelemetry } from "../../middleware/addStartAndEndTelemetry";
import { getSPOTenant } from "./deployDriver";
import { SPOClient } from "./utility/spoClient";
import {
  getPreparedSharePointProductId,
  PreparedPublicationError,
  PreparedPublicationOutput,
  PreparedPublicationTenantContext,
  isPreparedCatalogUrl,
  preparePublication,
} from "./utility/preparedPublication";
import { SharePointPackageInspectionError } from "./utility/packageFile";

export const preparePublicationActionName = "spfx/preparePublication";

type OutputKey = keyof PreparedPublicationOutput;

const outputKeys: OutputKey[] = [
  "publicationId",
  "preparedPublicationPath",
  "sourcePackageSha256",
  "sourceAppPackageSha256",
  "pluginSha256",
  "sharePointPackagePath",
  "sharePointPackageSha256",
  "appPackagePath",
  "appPackageSha256",
];

function isOutputKey(value: string): value is OutputKey {
  return outputKeys.some((key) => key === value);
}

interface PreparePublicationArgs {
  packagePath: string;
  appPackagePath: string;
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function validateArgs(value: unknown): PreparePublicationArgs | undefined {
  if (
    !isRecord(value) ||
    typeof value["packagePath"] !== "string" ||
    value["packagePath"].length === 0 ||
    typeof value["appPackagePath"] !== "string" ||
    value["appPackagePath"].length === 0
  ) {
    return undefined;
  }
  const unexpected = Object.keys(value).filter(
    (key) => key !== "packagePath" && key !== "appPackagePath"
  );
  return unexpected.length === 0
    ? { packagePath: value["packagePath"], appPackagePath: value["appPackagePath"] }
    : undefined;
}

function getClaims(value: Record<string, unknown>): Record<string, unknown> {
  const nested = value["value"];
  return isRecord(nested) ? nested : value;
}

function isGraphAudience(value: unknown): boolean {
  return (
    value === "https://graph.microsoft.com" || value === "00000003-0000-0000-c000-000000000000"
  );
}

function isSharePointAudience(value: unknown, tenantUrl: string): boolean {
  if (
    value === "00000003-0000-0ff1-ce00-000000000000" ||
    value === `${new URL(tenantUrl).origin}/`
  ) {
    return true;
  }
  return value === new URL(tenantUrl).origin;
}

async function discoverTenantContext(
  args: PreparePublicationArgs,
  context: DriverContext
): Promise<PreparedPublicationTenantContext> {
  const productId = await getPreparedSharePointProductId(context.projectPath, args.packagePath);
  if (productId.isErr()) {
    throw productId.error;
  }
  if (context.signal?.aborted) {
    throw new PreparedPublicationError(
      "PublicationCancelled",
      "Prepared publication was cancelled before tenant discovery."
    );
  }
  const tenantUrl = await getSPOTenant(context.m365TokenProvider);
  let tenantOrigin: string;
  try {
    const parsedTenantUrl = new URL(tenantUrl);
    if (parsedTenantUrl.protocol !== "https:" || !parsedTenantUrl.hostname) {
      throw new Error("Unsupported tenant URL.");
    }
    tenantOrigin = parsedTenantUrl.origin;
  } catch (error) {
    throw new PreparedPublicationError(
      "InvalidPreparedTenantContext",
      "The tenant URL could not be verified.",
      { cause: error }
    );
  }
  const graphClaimsResult = await context.m365TokenProvider.getJsonObject({
    scopes: GraphScopes,
  });
  if (graphClaimsResult.isErr()) {
    throw new PreparedPublicationError(
      "GraphTokenIdentityUnavailable",
      "The Graph token identity could not be verified.",
      { cause: graphClaimsResult.error }
    );
  }
  const graphClaims = getClaims(graphClaimsResult.value);
  const graphTenantId = graphClaims["tid"];
  if (
    typeof graphTenantId !== "string" ||
    !/^[0-9a-f]{8}-[0-9a-f]{4}-[1-8][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i.test(
      graphTenantId
    ) ||
    !isGraphAudience(graphClaims["aud"])
  ) {
    throw new PreparedPublicationError(
      "GraphTokenIdentityMismatch",
      "The Graph token tenant or audience does not match the prepared publication."
    );
  }
  const sharePointToken = await getSPFxToken(context.m365TokenProvider);
  if (!sharePointToken) {
    throw new PreparedPublicationError(
      "SharePointAuthenticationUnavailable",
      "A SharePoint access token is unavailable."
    );
  }
  const sharePointClaimsResult = await context.m365TokenProvider.getJsonObject({
    scopes: SPFxScopes(tenantOrigin),
  });
  if (sharePointClaimsResult.isErr()) {
    throw new PreparedPublicationError(
      "SharePointTokenIdentityUnavailable",
      "The SharePoint token identity could not be verified.",
      { cause: sharePointClaimsResult.error }
    );
  }
  const sharePointClaims = getClaims(sharePointClaimsResult.value);
  if (
    sharePointClaims["tid"] !== graphTenantId ||
    !isSharePointAudience(sharePointClaims["aud"], tenantOrigin)
  ) {
    throw new PreparedPublicationError(
      "SharePointTokenIdentityMismatch",
      "The SharePoint token tenant or audience does not match the prepared publication."
    );
  }
  if (context.signal?.aborted) {
    throw new PreparedPublicationError(
      "PublicationCancelled",
      "Prepared publication was cancelled before catalog discovery."
    );
  }
  const catalogUrl = await SPOClient.getAppCatalogSiteForTenant(
    sharePointToken,
    tenantOrigin,
    context.signal
  );
  if (!catalogUrl || !isPreparedCatalogUrl(catalogUrl, tenantOrigin)) {
    throw new PreparedPublicationError(
      "TenantAppCatalogUnavailable",
      "A supported tenant app catalog could not be discovered."
    );
  }
  return {
    tenantId: graphTenantId,
    tenantUrl: tenantOrigin,
    catalogUrl,
    productId: productId.value,
  };
}

function toFxError(error: PreparedPublicationError | SharePointPackageInspectionError): FxError {
  const message = getLocalizedString("driver.spfx.prepared.error", error.message);
  if (
    error instanceof PreparedPublicationError &&
    (error.code.startsWith("Journal") ||
      error.code.startsWith("PublicationWrite") ||
      error.code.startsWith("PublicationPrepare") ||
      error.code.startsWith("ProjectPath"))
  ) {
    const options: SystemErrorOptions = {
      source: preparePublicationActionName,
      error,
      message: error.code,
      displayMessage: message,
      categories: [ErrorCategory.External],
    };
    return new SystemError(options);
  }
  const options: UserErrorOptions = {
    source: preparePublicationActionName,
    error,
    message: error.code,
    displayMessage: message,
    categories: [ErrorCategory.External],
  };
  return new UserError(options);
}

@Service(preparePublicationActionName)
export class PreparePublicationDriver implements StepDriver {
  public readonly description = getLocalizedString("driver.spfx.prepared.description");
  public readonly progressTitle = getLocalizedString("driver.spfx.prepared.progress");

  @hooks([addStartAndEndTelemetry(preparePublicationActionName, preparePublicationActionName)])
  public async execute(
    args: unknown,
    context: DriverContext,
    outputEnvVarNames?: Map<string, string>
  ): Promise<ExecutionResult> {
    const validatedArgs = validateArgs(args);
    if (!validatedArgs) {
      return {
        result: err(
          new UserError({
            source: preparePublicationActionName,
            message: "InvalidActionInput",
            displayMessage: getLocalizedString(
              "driver.spfx.prepared.invalidInput",
              "packagePath, appPackagePath"
            ),
            categories: [ErrorCategory.External],
          })
        ),
        summaries: [],
      };
    }

    if (
      outputEnvVarNames === undefined ||
      outputEnvVarNames.size !== outputKeys.length ||
      outputKeys.some((key) => !outputEnvVarNames.has(key)) ||
      new Set(outputEnvVarNames.values()).size !== outputEnvVarNames.size ||
      [...outputEnvVarNames].some(
        ([key, variable]) => !isOutputKey(key) || !/^[A-Za-z_][A-Za-z0-9_]*$/.test(variable)
      )
    ) {
      return {
        result: err(
          new UserError({
            source: preparePublicationActionName,
            message: "InvalidActionOutput",
            displayMessage: getLocalizedString(
              "driver.spfx.prepared.invalidOutput",
              "writeToEnvironmentFile"
            ),
            categories: [ErrorCategory.External],
          })
        ),
        summaries: [],
      };
    }

    let result;
    try {
      const tenantContext = await discoverTenantContext(validatedArgs, context);
      result = await preparePublication({
        projectPath: context.projectPath,
        packagePath: validatedArgs.packagePath,
        appPackagePath: validatedArgs.appPackagePath,
        tenantContext,
        signal: context.signal,
      });
    } catch (error) {
      const preparedError =
        error instanceof PreparedPublicationError ||
        error instanceof SharePointPackageInspectionError
          ? error
          : new PreparedPublicationError(
              "TenantContextDiscoveryFailed",
              "The tenant publication context could not be verified.",
              { cause: error }
            );
      const fxError = toFxError(preparedError);
      context.logProvider.error(maskSecret(fxError.message));
      return { result: err(fxError), summaries: [] };
    }
    if (result.isErr()) {
      const error = toFxError(result.error);
      context.logProvider.error(maskSecret(error.message));
      return { result: err(error), summaries: [] };
    }

    const values = result.value;
    const outputs = new Map<string, string>();
    for (const key of outputKeys) {
      const value = values[key];
      const variable = outputEnvVarNames.get(key);
      if (variable !== undefined) {
        outputs.set(variable, value);
      }
    }
    return {
      result: ok(outputs),
      summaries: [getLocalizedString("driver.spfx.prepared.summary")],
    };
  }
}
