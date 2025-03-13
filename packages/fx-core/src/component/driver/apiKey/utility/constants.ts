// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

export const logMessageKeys = {
  startExecuteDriver: "driver.apiKey.log.startExecuteDriver",
  skipCreateApiKey: "driver.apiKey.log.skipCreateApiKey",
  apiKeyNotFound: "driver.apiKey.log.apiKeyNotFound",
  successCreateApiKey: "driver.apiKey.log.successCreateApiKey",
  failedExecuteDriver: "driver.apiKey.log.failedExecuteDriver",
  skipUpdateApiKey: "driver.apiKey.log.skipUpdateApiKey",
  successUpdateApiKey: "driver.apiKey.log.successUpdateApiKey",
};

export const maxDomainPerApiKey = 1;
export const maxSecretLength = 512;
export const minSecretLength = 10;

export const telemetryKeys = {
  isCustomAPIKey: "is-custom-api-key",
};
