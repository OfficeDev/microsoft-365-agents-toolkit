// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { TokenCredential } from "@azure/core-auth";
import { AccessToken, GetTokenOptions } from "@azure/identity";
import {
  AzureAccountProvider,
  Colors,
  ConfirmConfig,
  ConfirmResult,
  CryptoProvider,
  FxError,
  InputResult,
  InputTextConfig,
  InputTextResult,
  IProgressHandler,
  LoginStatus,
  LogLevel,
  LogProvider,
  M365TokenProvider,
  MultiSelectConfig,
  MultiSelectResult,
  ok,
  OptionItem,
  Result,
  SelectFileConfig,
  SelectFileResult,
  SelectFilesConfig,
  SelectFilesResult,
  SelectFolderConfig,
  SelectFolderResult,
  SingleFileOrInputConfig,
  SingleSelectConfig,
  SingleSelectResult,
  SubscriptionInfo,
  TelemetryReporter,
  TokenProvider,
  TokenRequest,
  Tools,
  UserInteraction,
} from "@microsoft/teamsfx-api";
import fs from "fs-extra";

export function randomAppName() {
  return `mock${Date.now()}-${Math.floor(Math.random() * 1000000)}`;
}

export class MyTokenCredential implements TokenCredential {
  public async getToken(
    scopes: string | string[],
    options?: GetTokenOptions
  ): Promise<AccessToken | null> {
    return {
      token: "a.eyJ1c2VySWQiOiJ0ZXN0QHRlc3QuY29tIn0=.c",
      expiresOnTimestamp: 1234,
    };
  }
}

export class MockedAzureAccountProvider implements AzureAccountProvider {
  async getIdentityCredentialAsync(showDialog?: boolean): Promise<TokenCredential | undefined> {
    return new MyTokenCredential();
  }

  async signout(): Promise<boolean> {
    return true;
  }
  async switchTenant(tenantId: string): Promise<Result<TokenCredential, FxError>> {
    return ok(new MyTokenCredential());
  }
  async setStatusChangeMap(
    name: string,
    statusChange: (
      status: string,
      token?: string,
      accountInfo?: Record<string, unknown>
    ) => Promise<void>,
    immediateCall?: boolean
  ): Promise<boolean> {
    return true;
  }
  async removeStatusChangeMap(name: string): Promise<boolean> {
    return true;
  }
  async getJsonObject(showDialog?: boolean): Promise<Record<string, unknown>> {
    return {};
  }
  async listSubscriptions(): Promise<SubscriptionInfo[]> {
    return [];
  }
  async setSubscription(subscriptionId: string): Promise<void> {}
  getAccountInfo(): Record<string, string> {
    return {};
  }
  async getSelectedSubscription(triggerUI?: boolean): Promise<SubscriptionInfo> {
    return {
      subscriptionId: "",
      subscriptionName: "",
      tenantId: "",
    };
  }
}

export class MockedM365Provider implements M365TokenProvider {
  async getAccessToken(tokenRequest: TokenRequest): Promise<Result<string, FxError>> {
    return ok("fakeToken");
  }
  async getJsonObject(
    tokenRequest: TokenRequest
  ): Promise<Result<Record<string, unknown>, FxError>> {
    return ok({
      upn: "fakeUserPrincipalName@fake.com",
      tid: "tenantId",
    });
  }
  async signout(): Promise<boolean> {
    return true;
  }
  async switchTenant(tenantId: string): Promise<Result<string, FxError>> {
    return ok("fakeToken");
  }
  async getStatus(tokenRequest: TokenRequest): Promise<Result<LoginStatus, FxError>> {
    return ok({
      status: "SignedIn",
      token: "fakeToken",
    });
  }
  async setStatusChangeMap(
    name: string,
    tokenRequest: TokenRequest,
    statusChange: (
      status: string,
      token?: string,
      accountInfo?: Record<string, unknown>
    ) => Promise<void>,
    immediateCall?: boolean
  ): Promise<Result<boolean, FxError>> {
    return ok(true);
  }
  async removeStatusChangeMap(name: string): Promise<Result<boolean, FxError>> {
    return ok(true);
  }
}

export class MockTelemetryReporter implements TelemetryReporter {
  sendTelemetryErrorEvent(
    eventName: string,
    properties?: { [key: string]: string },
    measurements?: { [key: string]: number },
    errorProps?: string[]
  ): void {
    // do nothing
  }

  sendTelemetryEvent(
    eventName: string,
    properties?: { [key: string]: string },
    measurements?: { [key: string]: number }
  ): void {
    // do nothing
  }

  sendTelemetryException(
    error: Error,
    properties?: { [key: string]: string },
    measurements?: { [key: string]: number }
  ): void {
    // do nothing
  }
}

export class MockUserInteraction implements UserInteraction {
  async selectOption(config: SingleSelectConfig): Promise<Result<SingleSelectResult, FxError>> {
    return ok({ type: "success" });
  }

  selectOptions(config: MultiSelectConfig): Promise<Result<MultiSelectResult, FxError>> {
    // Return the first option as selected by default for tests
    const options = config.options as (string | OptionItem)[];
    const result =
      options?.length > 0 ? [typeof options[0] === "string" ? options[0] : options[0].id] : [];
    return Promise.resolve(ok({ type: "success", result }));
  }

  inputText(config: InputTextConfig): Promise<Result<InputTextResult, FxError>> {
    throw new Error(`Method inputText not implemented: ${JSON.stringify(config)}`);
  }

  selectFile(config: SelectFileConfig): Promise<Result<SelectFileResult, FxError>> {
    throw new Error(`Method selectFile not implemented: ${JSON.stringify(config)}`);
  }

  selectFiles(config: SelectFilesConfig): Promise<Result<SelectFilesResult, FxError>> {
    throw new Error(`Method selectFiles not implemented: ${JSON.stringify(config)}`);
  }

  selectFolder(config: SelectFolderConfig): Promise<Result<SelectFolderResult, FxError>> {
    throw new Error(`Method selectFolder not implemented: ${JSON.stringify(config)}`);
  }

  openUrl(link: string): Promise<Result<boolean, FxError>> {
    throw new Error(`Method openUrl not implemented: ${link}`);
  }

  selectFileOrInput(
    config: SingleFileOrInputConfig
  ): Promise<Result<InputResult<string>, FxError>> {
    throw new Error(`Method selectFileOrInput not implemented: ${config}`);
  }

  async showMessage(
    level: "info" | "warn" | "error",
    message: string,
    modal: boolean,
    ...items: string[]
  ): Promise<Result<string | undefined, FxError>>;

  async showMessage(
    level: "info" | "warn" | "error",
    message: Array<{ content: string; color: Colors }>,
    modal: boolean,
    ...items: string[]
  ): Promise<Result<string | undefined, FxError>>;

  async showMessage(
    level: "info" | "warn" | "error",
    message: string | Array<{ content: string; color: Colors }>,
    modal: boolean,
    ...items: string[]
  ): Promise<Result<string | undefined, FxError>> {
    return ok("");
  }

  createProgressBar(title: string, totalSteps: number): IProgressHandler {
    const handler: IProgressHandler = {
      start: async (detail?: string): Promise<void> => {},
      next: async (detail?: string): Promise<void> => {},
      end: async (): Promise<void> => {},
    };
    return handler;
  }

  async runCommand(args: {
    cmd: string;
    workingDirectory?: string;
    shell?: string;
    timeout?: number;
    env?: { [k: string]: string };
  }): Promise<Result<string, FxError>> {
    return ok("");
  }

  async confirm(config: ConfirmConfig): Promise<Result<ConfirmResult, FxError>> {
    return ok({ type: "success", result: true });
  }

  async openFile(filePath: string): Promise<Result<boolean, FxError>> {
    return ok(true);
  }
}

export class MockTools implements Tools {
  logProvider = new MockLogProvider();
  tokenProvider: TokenProvider = {
    azureAccountProvider: new MockedAzureAccountProvider(),
    m365TokenProvider: new MockedM365Provider(),
  };
  telemetryReporter = new MockTelemetryReporter();
  ui = new MockUserInteraction();
  cryptoProvider = new MockCryptoProvider();
}

export class MockCryptoProvider implements CryptoProvider {
  encrypt(plaintext: string): Result<string, FxError> {
    return ok(plaintext);
  }

  decrypt(ciphertext: string): Result<string, FxError> {
    return ok(ciphertext);
  }
}

export class MockLogProvider implements LogProvider {
  msg = "";
  verbose(msg: string): void {
    this.log(LogLevel.Verbose, msg);
  }
  debug(msg: string): void {
    this.log(LogLevel.Debug, msg);
  }
  info(msg: string | Array<any>): void {
    this.log(LogLevel.Info, msg as string);
  }
  warning(msg: string): void {
    this.log(LogLevel.Warning, msg);
  }
  error(msg: string): void {
    this.log(LogLevel.Error, msg);
  }
  log(level: LogLevel, msg: string): void {
    this.msg = msg;
  }
  async logInFile(level: LogLevel, msg: string): Promise<void> {
    this.msg = msg;
  }
  getLogFilePath(): string {
    return "";
  }
}

export function deleteFolder(filePath?: string): void {
  if (!filePath) return;
  if (fs.existsSync(filePath)) {
    const files = fs.readdirSync(filePath);
    files.forEach((file) => {
      const nextFilePath = `${filePath}/${file}`;
      const states = fs.statSync(nextFilePath);
      if (states.isDirectory()) {
        //recurse
        deleteFolder(nextFilePath);
      } else {
        //delete file
        fs.unlinkSync(nextFilePath);
      }
    });
    try {
      fs.rmdirSync(filePath);
    } catch (e) {}
  }
}
