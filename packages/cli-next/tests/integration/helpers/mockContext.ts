/**
 * Copyright (c) Microsoft Corporation.
 * Licensed under the MIT license.
 *
 * Shared mock AtkContext for CLI integration tests.
 * Stubs auth (returns fake tokens), telemetry, UI (auto-confirms),
 * and logger (silent).
 */

import * as sinon from "sinon";
import * as os from "node:os";
import * as path from "node:path";
import * as fs from "node:fs/promises";
import { ok } from "neverthrow";
import type { AtkContext } from "@microsoft/teamsfx-core-next";

const MOCK_M365_TOKEN = "mock-m365-token-for-testing";

/**
 * Create a fully-stubbed AtkContext suitable for lifecycle integration tests.
 *
 * - Auth: returns fake tokens without real login
 * - UI: auto-confirms all prompts, selects first option
 * - Logger: silent (sinon stubs)
 * - Telemetry: captured in stubs for assertions
 */
export function createIntegrationContext(projectPath: string): AtkContext {
  return {
    auth: {
      m365TokenProvider: {
        getAccessToken: sinon.stub().resolves(ok({ token: MOCK_M365_TOKEN })),
        getJsonObject: sinon.stub().resolves(ok({ tid: "mock-tenant-id" })),
        getStatus: sinon.stub().resolves(ok({ status: "SignedIn", accountInfo: { name: "Test" } })),
        setStatusChangeMap: sinon.stub(),
        removeStatusChangeMap: sinon.stub(),
      } as any,
      azureAccountProvider: {
        getIdentityCredentialAsync: sinon.stub().resolves({ getToken: sinon.stub() }),
        signout: sinon.stub().resolves(true),
        setStatusChangeMap: sinon.stub(),
        removeStatusChangeMap: sinon.stub(),
        getJsonObject: sinon.stub().resolves(ok({ tid: "mock-azure-tenant" })),
        listSubscriptions: sinon.stub().resolves([
          {
            subscriptionId: "00000000-0000-0000-0000-000000000000",
            subscriptionName: "Test Sub",
            tenantId: "mock-azure-tenant",
          },
        ]),
        setSubscription: sinon.stub().resolves(),
        getAccountInfo: sinon
          .stub()
          .returns({ subscriptionId: "00000000-0000-0000-0000-000000000000" }),
        getSelectedSubscription: sinon.stub().resolves({
          subscriptionId: "00000000-0000-0000-0000-000000000000",
          subscriptionName: "Test Sub",
          tenantId: "mock-azure-tenant",
        }),
      } as any,
    },
    logger: {
      log: sinon.stub(),
      verbose: sinon.stub(),
      debug: sinon.stub(),
      info: sinon.stub(),
      warning: sinon.stub(),
      error: sinon.stub(),
      logInFile: sinon.stub().resolves(),
      getLogFilePath: sinon.stub().returns("/dev/null"),
    } as any,
    telemetry: {
      sendTelemetryEvent: sinon.stub(),
      sendTelemetryErrorEvent: sinon.stub(),
      sendTelemetryException: sinon.stub(),
    },
    ui: {
      selectOption: sinon.stub().resolves(ok({ type: "success", result: "" })),
      selectOptions: sinon.stub().resolves(ok({ type: "success", result: [] })),
      inputText: sinon.stub().resolves(ok({ type: "success", result: "test-input" })),
      selectFile: sinon.stub().resolves(ok({ type: "success", result: "" })),
      selectFiles: sinon.stub().resolves(ok({ type: "success", result: [] })),
      selectFolder: sinon.stub().resolves(ok({ type: "success", result: "" })),
      openUrl: sinon.stub().resolves(ok(true)),
      showMessage: sinon.stub().resolves(ok("")),
      createProgressBar: sinon.stub().returns({
        start: sinon.stub().resolves(),
        next: sinon.stub().resolves(),
        end: sinon.stub().resolves(),
      }),
      confirm: sinon.stub().resolves(ok({ type: "success", result: true })),
    } as any,
    correlationId: "integration-test-" + Math.random().toString(36).slice(2, 10),
    projectPath,
  };
}

/**
 * Create a temporary directory for a test and return its path.
 * Caller must clean up with `fs.rm(dir, { recursive: true, force: true })`.
 */
export async function makeTmpDir(prefix = "cli-pipeline-"): Promise<string> {
  return fs.mkdtemp(path.join(os.tmpdir(), prefix));
}
