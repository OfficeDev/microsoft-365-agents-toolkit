// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

class EnvironmentNameManager {
  public readonly envNameRegex = /^[\w\d-_]+$/;
  public readonly envConfigNameRegex = /^config\.(?<envName>[\w\d-_]+)\.json$/i;
  public readonly envStateNameRegex = /^state\.(?<envName>[\w\d-_]+)\.json$/i;

  public readonly schema = "https://aka.ms/teamsfx-env-config-schema";

  private readonly defaultEnvName = "dev";
  private readonly localEnvName = "local";
  private readonly testToolEnvName = "testtool";
  private readonly sandboxEnvName = "sandbox";

  public getDefaultEnvName() {
    return this.defaultEnvName;
  }

  public getLocalEnvName() {
    return this.localEnvName;
  }

  public getTestToolEnvName() {
    return this.testToolEnvName;
  }

  public getSandboxEnvName() {
    return this.sandboxEnvName;
  }

  public getNonRemoteEnvNames(): string[] {
    return [this.localEnvName, this.testToolEnvName];
  }

  public isRemoteEnvironment(env: string) {
    return this.getNonRemoteEnvNames().indexOf(env) === -1;
  }
}

export const environmentNameManager = new EnvironmentNameManager();
