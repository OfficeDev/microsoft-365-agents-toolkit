/**
 * Copyright (c) Microsoft Corporation.
 * Licensed under the MIT license.
 */

import * as sinon from "sinon";
import { AtkContext } from "../../src/core/context";

/**
 * Create a fully mocked AtkContext for unit testing.
 * All methods are Sinon stubs that can be inspected.
 */
export function createMockContext(overrides?: Partial<AtkContext>): AtkContext & {
  telemetry: sinon.SinonStubbedInstance<AtkContext["telemetry"]>;
  logger: sinon.SinonStubbedInstance<AtkContext["logger"]>;
} {
  const ctx: AtkContext = {
    auth: {
      m365TokenProvider: {} as any,
      azureAccountProvider: {} as any,
    },
    logger: {
      log: sinon.stub(),
      verbose: sinon.stub(),
      debug: sinon.stub(),
      info: sinon.stub(),
      warning: sinon.stub(),
      error: sinon.stub(),
      logInFile: sinon.stub().resolves(),
      getLogFilePath: sinon.stub().returns("/tmp/test.log"),
    } as any,
    telemetry: {
      sendTelemetryEvent: sinon.stub(),
      sendTelemetryErrorEvent: sinon.stub(),
      sendTelemetryException: sinon.stub(),
    },
    ui: {
      selectOption: sinon.stub(),
      selectOptions: sinon.stub(),
      inputText: sinon.stub(),
      selectFile: sinon.stub(),
      selectFiles: sinon.stub(),
      selectFolder: sinon.stub(),
      openUrl: sinon.stub(),
      showMessage: sinon.stub(),
      createProgressBar: sinon.stub(),
      confirm: sinon.stub(),
    } as any,
    correlationId: "test-correlation-id-1234",
    projectPath: "/tmp/test-project",
    ...overrides,
  };
  return ctx as any;
}
