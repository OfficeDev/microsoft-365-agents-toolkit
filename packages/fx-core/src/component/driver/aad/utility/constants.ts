// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

export const logMessageKeys = {
  startExecuteDriver: "driver.aadApp.log.startExecuteDriver",
  successExecuteDriver: "driver.aadApp.log.successExecuteDriver",
  failExecuteDriver: "driver.aadApp.log.failExecuteDriver",
  startCreateAadApp: "driver.aadApp.log.startCreateAadApp",
  successCreateAadApp: "driver.aadApp.log.successCreateAadApp",
  skipCreateAadApp: "driver.aadApp.log.skipCreateAadApp",
  startGenerateClientSecret: "driver.aadApp.log.startGenerateClientSecret",
  successGenerateClientSecret: "driver.aadApp.log.successGenerateClientSecret",
  skipGenerateClientSecret: "driver.aadApp.log.skipGenerateClientSecret",
  outputAadAppManifest: "driver.aadApp.log.outputAadAppManifest",
  successUpdateAadAppManifest: "driver.aadApp.log.successUpdateAadAppManifest",
  deleteAadAfterDebugging: "driver.aadApp.log.deleteAadAfterDebugging",
  insufficientPermission: "driver.aadApp.log.insufficientPermission",
};

export const descriptionMessageKeys = {
  create: "driver.aadApp.description.create",
  update: "driver.aadApp.description.update",
};

export const questionKeys = {
  aadAppIdTitle: "driver.aadApp.question.id.title",
  addAppIdValidation: "driver.aadApp.question.id.validation",
  aadAppSecretTitle: "driver.aadApp.question.secret.title",
  aadAppSecretValidation: "driver.aadApp.question.secret.validation",
  aadAppObjectIdTitle: "driver.aadApp.question.objectId.title",
  aadAppObjectIdValidation: "driver.aadApp.question.objectId.validation",
};

export const permissionsKeys = {
  name: "Microsoft Entra App",
  owner: "Owner",
  noPermission: "No Permission",
  type: "M365",
};

export const aadErrorCode = {
  permissionErrorCode: "CannotDeleteOrUpdateEnabledEntitlement",
  hostNameNotOnVerifiedDomain: "HostNameNotOnVerifiedDomain", // Using unverified domain in multi tenant scenario
  credentialInvalidLifetimeAsPerAppPolicy: "CredentialInvalidLifetimeAsPerAppPolicy",
  credentialTypeNotAllowedAsPerAppPolicy: "CredentialTypeNotAllowedAsPerAppPolicy",
};

export const constants = {
  aadAppPasswordDisplayName: "default",
  oauthAuthorityPrefix: "https://login.microsoftonline.com",
  defaultHelpLink: "https://aka.ms/teamsfx-actions/aadapp-create",
  missingServiceManagementReferenceHelpLink:
    "https://aka.ms/teamsfx/missing-service-management-reference-help",
  insufficientPermissionErrorMessage: "Insufficient privileges to complete the operation.",
};

export const telemetryKeys = {
  newAadApp: "new-aad-app",
  userInputAadApp: "user-input-aad-app",
  insufficientPermissionAadApp: "insufficient-permission-aad-app",
};
