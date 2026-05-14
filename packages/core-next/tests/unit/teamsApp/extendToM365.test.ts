/**
 * Copyright (c) Microsoft Corporation.
 * Licensed under the MIT license.
 */

import { expect } from "chai";
import { describe, it, afterEach } from "mocha";
import * as sinon from "sinon";
import * as fs from "node:fs/promises";
import * as path from "node:path";
import * as os from "node:os";
import { ok, err } from "neverthrow";
import { runOperation } from "../../../src/core/operation";
import { extendToM365Op } from "../../../src/teamsApp/operations";
import { createMockContext } from "../testHelper";

describe("teamsApp/extendToM365Op", () => {
  const sandbox = sinon.createSandbox();
  let tmpDir: string;

  afterEach(async () => {
    sandbox.restore();
    if (tmpDir) {
      await fs.rm(tmpDir, { recursive: true, force: true });
    }
  });

  async function setup(): Promise<string> {
    tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), "atk-test-m365-"));
    return tmpDir;
  }

  it("should return error when app package not found", async () => {
    const dir = await setup();
    const ctx = createMockContext({ projectPath: dir });

    const result = await runOperation(extendToM365Op, ctx, {
      appPackagePath: path.join(dir, "nonexistent.zip"),
    });

    expect(result.isErr()).to.be.true;
    expect(result._unsafeUnwrapErr().code).to.equal("AppPackageNotFound");
  });

  it("should return error when token acquisition fails", async () => {
    const dir = await setup();
    const packagePath = path.join(dir, "app.zip");
    await fs.writeFile(packagePath, "fake-zip");

    const ctx = createMockContext({
      projectPath: dir,
      auth: {
        m365TokenProvider: {
          getAccessToken: sandbox.stub().resolves(err(new Error("Token failed"))),
          getJsonObject: sandbox.stub(),
          getStatus: sandbox.stub(),
          setStatusChangeMap: sandbox.stub(),
          removeStatusChangeMap: sandbox.stub(),
          signout: sandbox.stub(),
        } as any,
        azureAccountProvider: {} as any,
      },
    });

    const result = await runOperation(extendToM365Op, ctx, {
      appPackagePath: packagePath,
    });

    expect(result.isErr()).to.be.true;
    expect(result._unsafeUnwrapErr().code).to.equal("TokenAcquisitionError");
  });

  it("should validate input schema rejects empty path", async () => {
    const ctx = createMockContext();

    const result = await runOperation(extendToM365Op, ctx, {
      appPackagePath: "",
    });

    expect(result.isErr()).to.be.true;
    // Zod validation error
    expect(result._unsafeUnwrapErr().code).to.equal("InputValidationError");
  });
});
