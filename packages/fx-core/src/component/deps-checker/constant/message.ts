// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { getDefaultString, getLocalizedString } from "../../../common/localizeUtils";
import { nodeInstallationLink } from "./helpLink";

export const Messages = {
  needInstallNpm: () => ({
    default: getDefaultString("depChecker.needInstallNpm"),
    localized: getLocalizedString("depChecker.needInstallNpm"),
  }),
  failToValidateFuncCoreTool: (functionVersionErrMsg: string) => ({
    default: getDefaultString("depChecker.failToValidateFuncCoreTool", functionVersionErrMsg),
    localized: getLocalizedString("depChecker.failToValidateFuncCoreTool", functionVersionErrMsg),
  }),
  portableFuncNodeNotMatched: (nodeVersion: string, funcVersion: string) => ({
    default: getDefaultString("depChecker.portableFuncNodeNotMatched")
      .replace("@NodeVersion", nodeVersion)
      .replace("@FuncVersion", funcVersion),
    localized: getLocalizedString("depChecker.portableFuncNodeNotMatched")
      .replace("@NodeVersion", nodeVersion)
      .replace("@FuncVersion", funcVersion),
  }),
  symlinkDirAlreadyExist: (linkFilePath: string) => ({
    default: getDefaultString("depChecker.symlinkDirAlreadyExist", linkFilePath),
    localized: getLocalizedString("depChecker.symlinkDirAlreadyExist", linkFilePath),
  }),
  invalidFuncVersion: (version: string) => ({
    default: getDefaultString("depChecker.invalidFuncVersion", version),
    localized: getLocalizedString("depChecker.invalidFuncVersion", version),
  }),
  noSentinelFile: () => ({
    default: getDefaultString("depChecker.noSentinelFile"),
    localized: getLocalizedString("depChecker.noSentinelFile"),
  }),
  funcVersionNotMatch: (funcVersion: string, expectedFuncVersion: string) => ({
    default: getDefaultString("depChecker.funcVersionNotMatch", funcVersion, expectedFuncVersion),
    localized: getLocalizedString(
      "depChecker.funcVersionNotMatch",
      funcVersion,
      expectedFuncVersion
    ),
  }),

  downloadDotnet: () => getLocalizedString("depChecker.downloadDotnet"),
  finishInstallDotnet: () => getLocalizedString("depChecker.finishInstallDotnet"),
  useGlobalDotnet: () => getLocalizedString("depChecker.useGlobalDotnet"),
  dotnetInstallStderr: () => getLocalizedString("depChecker.dotnetInstallStderr"),
  dotnetInstallErrorCode: () => getLocalizedString("depChecker.dotnetInstallErrorCode"),

  dotnetNotFound: () => getLocalizedString("depChecker.dotnetNotFound"),
  // depsNotFound: () => getLocalizedString("depChecker.depsNotFound"),

  testToolVersionNotMatch: (version: string, expectedVersion: string) => ({
    default: getDefaultString("depChecker.testToolVersionNotMatch", version, expectedVersion),
    localized: getLocalizedString("depChecker.testToolVersionNotMatch", version, expectedVersion),
  }),
  failToValidateTestTool: (errorMessage: string) => ({
    default: getDefaultString("depChecker.failedToValidateTestTool", errorMessage),
    localized: getLocalizedString("depChecker.failedToValidateTestTool", errorMessage),
  }),

  failToValidateVxTestAppInstallOptions: () => ({
    default: getDefaultString("depChecker.failToValidateVxTestAppInstallOptions"),
    localized: getLocalizedString("depChecker.failToValidateVxTestAppInstallOptions"),
  }),
  failToValidateVxTestApp: () => ({
    default: getDefaultString("depChecker.failToValidateVxTestApp"),
    localized: getLocalizedString("depChecker.failToValidateVxTestApp"),
  }),

  failToDownloadFromUrl: () => getLocalizedString("depChecker.failToDownloadFromUrl"),

  linuxDepsNotFound: () => getLocalizedString("depChecker.linuxDepsNotFound"),
};
