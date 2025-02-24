// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { AADApplication } from "../interface/AADApplication";
import { AADManifest } from "../interface/AADManifest";
import isUUID from "validator/lib/isUUID";
import { getPermissionMap } from "../permissions";
import {
  AadManifestErrorMessage,
  MissingResourceAccessIdUserError,
  MissingResourceAppIdUserError,
  RequiredResourceAccessShouldBeArrayUserError,
  ResourceAccessShouldBeArrayUserError,
  UnknownResourceAccessIdUserError,
  UnknownResourceAccessTypeUserError,
  UnknownResourceAppIdUserError,
} from "../error/aadManifestError";
import { TOOLS } from "../../../../common/globalVars";
import { getLocalizedString } from "../../../../common/localizeUtils";
import { err, FxError, ok, Result } from "@microsoft/teamsfx-api";
import { FileNotFoundError, UserCancelError } from "../../../../error";
import fs from "fs-extra";
import { updateVersionForTeamsAppYamlFile } from "../../util/utils";

const componentName = "AadManifestHelper";

export class AadManifestHelper {
  public static manifestToApplication(manifest: AADManifest | AADApplication): AADApplication {
    if (this.isNewAADManifestSchema(manifest)) {
      return manifest as AADApplication;
    }

    manifest = manifest as AADManifest;

    const result: AADApplication = {
      id: manifest.id,
      appId: manifest.appId,
      disabledByMicrosoftStatus: manifest.disabledByMicrosoftStatus,
      displayName: manifest.name,
      description: manifest.description,
      groupMembershipClaims: manifest.groupMembershipClaims,
      identifierUris: manifest.identifierUris,
      isFallbackPublicClient: manifest.allowPublicClient,
      notes: manifest.notes,
      signInAudience: manifest.signInAudience,
      tags: manifest.tags,
      tokenEncryptionKeyId: manifest.tokenEncryptionKeyId,
      addIns: manifest.addIns,
      api: {
        acceptMappedClaims: manifest.acceptMappedClaims,
        knownClientApplications: manifest.knownClientApplications,
        requestedAccessTokenVersion: manifest.accessTokenAcceptedVersion,
        oauth2PermissionScopes: manifest.oauth2Permissions,
        preAuthorizedApplications: manifest.preAuthorizedApplications?.map((item) => {
          return { appId: item.appId, delegatedPermissionIds: item.permissionIds };
        }),
      },
      appRoles: manifest.appRoles,
      info: {
        marketingUrl: manifest.informationalUrls?.marketing,
        privacyStatementUrl: manifest.informationalUrls?.privacy,
        supportUrl: manifest.informationalUrls?.support,
        termsOfServiceUrl: manifest.informationalUrls?.termsOfService,
      },
      keyCredentials: manifest.keyCredentials?.map((item) => {
        return {
          customKeyIdentifier: item.customKeyIdentifier,
          displayName: item.displayName,
          endDateTime: item.endDate,
          key: item.value,
          keyId: item.keyId,
          startDateTime: item.startDate,
          type: item.type,
          usage: item.usage,
        };
      }),
      optionalClaims: manifest.optionalClaims,
      parentalControlSettings: manifest.parentalControlSettings,
      publicClient: {
        redirectUris: manifest.replyUrlsWithType
          ?.filter((item) => item.type === "InstalledClient")
          .map((item) => item.url),
      },
      requiredResourceAccess: manifest.requiredResourceAccess,
      web: {
        homePageUrl: manifest.signInUrl,
        logoutUrl: manifest.logoutUrl,
        redirectUris: manifest.replyUrlsWithType
          ?.filter((item) => item.type === "Web")
          .map((item) => item.url),
        implicitGrantSettings: {
          enableIdTokenIssuance: manifest.oauth2AllowIdTokenImplicitFlow,
          enableAccessTokenIssuance: manifest.oauth2AllowImplicitFlow,
        },
      },
      spa: {
        redirectUris: manifest.replyUrlsWithType
          ?.filter((item) => item.type === "Spa")
          .map((item) => item.url),
      },
    };

    return result;
  }

  public static applicationToManifest(app: AADApplication): AADManifest {
    const result: AADManifest = {
      id: app.id,
      appId: app.appId,
      acceptMappedClaims: app.api.acceptMappedClaims,
      accessTokenAcceptedVersion: app.api.requestedAccessTokenVersion,
      addIns: app.addIns,
      allowPublicClient: app.isFallbackPublicClient,
      appRoles: app.appRoles,
      description: app.description,
      disabledByMicrosoftStatus: app.disabledByMicrosoftStatus,
      groupMembershipClaims: app.groupMembershipClaims,
      identifierUris: app.identifierUris,
      informationalUrls: {
        termsOfService: app.info.termsOfServiceUrl,
        support: app.info.supportUrl,
        privacy: app.info.privacyStatementUrl,
        marketing: app.info.marketingUrl,
      },
      keyCredentials: app.keyCredentials.map((item) => {
        return {
          customKeyIdentifier: item.customKeyIdentifier,
          endDate: item.endDateTime,
          keyId: item.keyId,
          startDate: item.startDateTime,
          type: item.type,
          usage: item.usage,
          value: item.key,
          displayName: item.displayName,
        };
      }),
      knownClientApplications: app.api.knownClientApplications,
      logoutUrl: app.web.logoutUrl,
      name: app.displayName,
      notes: app.notes,
      oauth2AllowIdTokenImplicitFlow: app.web.implicitGrantSettings.enableIdTokenIssuance,
      oauth2AllowImplicitFlow: app.web.implicitGrantSettings.enableIdTokenIssuance,
      oauth2Permissions: app.api.oauth2PermissionScopes,
      optionalClaims: app.optionalClaims,
      parentalControlSettings: app.parentalControlSettings,
      preAuthorizedApplications: app.api.preAuthorizedApplications.map((item) => {
        return {
          appId: item.appId,
          permissionIds: item.delegatedPermissionIds,
        };
      }),
      replyUrlsWithType: app.spa.redirectUris
        .map((item) => {
          return {
            type: "Spa",
            url: item,
          };
        })
        .concat(
          app.web.redirectUris.map((item) => {
            return {
              type: "Web",
              url: item,
            };
          })
        )
        .concat(
          app.publicClient.redirectUris.map((item) => {
            return {
              type: "InstalledClient",
              url: item,
            };
          })
        ),
      requiredResourceAccess: app.requiredResourceAccess,
      signInUrl: app.web.homePageUrl,
      signInAudience: app.signInAudience,
      tags: app.tags,
      tokenEncryptionKeyId: app.tokenEncryptionKeyId,
    };

    return result;
  }

  public static validateManifest(manifest: AADManifest | AADApplication): string {
    let warningMsg = "";

    // if manifest doesn't contain signInAudience property
    if (!manifest.signInAudience) {
      warningMsg += AadManifestErrorMessage.SignInAudienceIsMissing;
    }

    // if manifest doesn't contain requiredResourceAccess property
    if (!manifest.requiredResourceAccess) {
      warningMsg += AadManifestErrorMessage.RequiredResourceAccessIsMissing;
    }

    if (AadManifestHelper.isNewAADManifestSchema(manifest)) {
      manifest = manifest as AADApplication;
      // if manifest doesn't contain name property or name is empty
      if (!manifest.displayName || manifest.displayName === "") {
        warningMsg += AadManifestErrorMessage.NameIsMissing;
      }

      // if manifest doesn't contain oauth2Permissions or oauth2Permissions length is 0
      if (
        !manifest.api?.oauth2PermissionScopes ||
        manifest.api?.oauth2PermissionScopes.length === 0
      ) {
        warningMsg += AadManifestErrorMessage.Oauth2PermissionsIsMissing;
      }

      // if manifest doesn't contain preAuthorizedApplications
      if (
        !manifest.api?.preAuthorizedApplications ||
        manifest.api?.preAuthorizedApplications.length === 0
      ) {
        warningMsg += AadManifestErrorMessage.PreAuthorizedApplicationsIsMissing;
      }

      // if accessTokenAcceptedVersion in manifest is not 2
      if (manifest.api?.requestedAccessTokenVersion !== 2) {
        warningMsg += AadManifestErrorMessage.AccessTokenAcceptedVersionIs1;
      }
    } else {
      manifest = manifest as AADManifest;
      // if manifest doesn't contain name property or name is empty
      if (!manifest.name || manifest.name === "") {
        warningMsg += AadManifestErrorMessage.NameIsMissing;
      }

      // if manifest doesn't contain oauth2Permissions or oauth2Permissions length is 0
      if (!manifest.oauth2Permissions || manifest.oauth2Permissions.length === 0) {
        warningMsg += AadManifestErrorMessage.Oauth2PermissionsIsMissing;
      }

      // if manifest doesn't contain preAuthorizedApplications
      if (!manifest.preAuthorizedApplications || manifest.preAuthorizedApplications.length === 0) {
        warningMsg += AadManifestErrorMessage.PreAuthorizedApplicationsIsMissing;
      }

      // if accessTokenAcceptedVersion in manifest is not 2
      if (manifest.accessTokenAcceptedVersion !== 2) {
        warningMsg += AadManifestErrorMessage.AccessTokenAcceptedVersionIs1;
      }
    }

    // if manifest doesn't contain optionalClaims or access token doesn't contain idtyp clams
    if (!manifest.optionalClaims) {
      warningMsg += AadManifestErrorMessage.OptionalClaimsIsMissing;
    } else if (!manifest.optionalClaims.accessToken?.find((item) => item.name === "idtyp")) {
      warningMsg += AadManifestErrorMessage.OptionalClaimsMissingIdtypClaim;
    }

    if (warningMsg) {
      warningMsg = AadManifestErrorMessage.AADManifestIssues + warningMsg;
    }
    return warningMsg.trimEnd();
  }

  public static processRequiredResourceAccessInManifest(
    manifest: AADManifest | AADApplication
  ): void {
    const map = getPermissionMap();

    if (manifest.requiredResourceAccess && !Array.isArray(manifest.requiredResourceAccess)) {
      throw new RequiredResourceAccessShouldBeArrayUserError(componentName);
    }

    manifest.requiredResourceAccess?.forEach((requiredResourceAccessItem) => {
      const resourceIdOrName = requiredResourceAccessItem.resourceAppId;
      let resourceId = resourceIdOrName;
      if (!resourceIdOrName) {
        throw new MissingResourceAppIdUserError(componentName);
      }
      if (!isUUID(resourceIdOrName)) {
        resourceId = map[resourceIdOrName]?.id;
        if (!resourceId) {
          throw new UnknownResourceAppIdUserError(componentName, resourceIdOrName);
        }
        requiredResourceAccessItem.resourceAppId = resourceId;
      }

      if (
        requiredResourceAccessItem.resourceAccess &&
        !Array.isArray(requiredResourceAccessItem.resourceAccess)
      ) {
        throw new ResourceAccessShouldBeArrayUserError(componentName);
      }

      requiredResourceAccessItem.resourceAccess?.forEach((resourceAccessItem) => {
        const resourceAccessIdOrName = resourceAccessItem.id;
        if (!resourceAccessIdOrName) {
          throw new MissingResourceAccessIdUserError(componentName);
        }
        if (!isUUID(resourceAccessIdOrName)) {
          let resourceAccessId;
          if (resourceAccessItem.type === "Scope") {
            resourceAccessId = map[resourceId]?.scopes[resourceAccessItem.id];
          } else if (resourceAccessItem.type === "Role") {
            resourceAccessId = map[resourceId]?.roles[resourceAccessItem.id];
          } else {
            throw new UnknownResourceAccessTypeUserError(componentName, resourceAccessItem.type);
          }

          if (!resourceAccessId) {
            throw new UnknownResourceAccessIdUserError(componentName, resourceAccessItem.id);
          }
          resourceAccessItem.id = resourceAccessId;
        }
      });
    });
  }

  public static isNewAADManifestSchema(manifest: AADManifest | AADApplication): boolean {
    return "displayName" in manifest;
  }

  public static async showWarningIfManifestIsOutdated(
    manifestTemplatePath: string,
    projectPath: string
  ): Promise<void> {
    const manifest = await fs.readJson(manifestTemplatePath);
    if (!AadManifestHelper.isNewAADManifestSchema(manifest)) {
      void TOOLS.ui
        .showMessage(
          "warn",
          getLocalizedString("core.convertAadToNewSchema.outdate"),
          false,
          getLocalizedString("core.convertAadToNewSchema.upgrade")
        )
        .then((result) => {
          if (
            result.isOk() &&
            result.value === getLocalizedString("core.convertAadToNewSchema.upgrade")
          ) {
            void AadManifestHelper.convertManifestToNewSchemaAndOverride(
              manifestTemplatePath,
              projectPath
            );
          }
        });
    }
  }

  public static async convertManifestToNewSchemaAndOverride(
    manifestTemplatePath: string,
    projectPath: string
  ): Promise<Result<undefined, FxError>> {
    if (!(await fs.pathExists(manifestTemplatePath))) {
      return err(new FileNotFoundError("convertAadToNewSchema", manifestTemplatePath));
    }
    const manifest = await fs.readJson(manifestTemplatePath);

    if (AadManifestHelper.isNewAADManifestSchema(manifest)) {
      void TOOLS.ui.showMessage(
        "info",
        getLocalizedString("core.convertAadToNewSchema.alreadyNewSchema"),
        false
      );
      return ok(undefined);
    }

    const confirmRes = await TOOLS.ui.showMessage(
      "warn",
      getLocalizedString("core.convertAadToNewSchema.warning"),
      true,
      getLocalizedString("core.convertAadToNewSchema.continue")
    );

    if (
      confirmRes.isOk() &&
      confirmRes.value !== getLocalizedString("core.convertAadToNewSchema.continue")
    ) {
      return err(new UserCancelError());
    }

    const result = AadManifestHelper.manifestToApplication(manifest);
    await fs.writeJson(manifestTemplatePath, result, { spaces: 2 });
    void TOOLS.ui.showMessage(
      "info",
      getLocalizedString("core.convertAadToNewSchema.success"),
      false
    );

    await updateVersionForTeamsAppYamlFile(projectPath);
    return ok(undefined);
  }
}
