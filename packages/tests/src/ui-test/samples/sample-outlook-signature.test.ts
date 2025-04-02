// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Ivan Chen <v-ivanchen@microsoft.com>
 */

import { TemplateProject, Timeout } from "../../utils/constants";
import { CaseFactory } from "./sampleCaseFactory";
import { SampledebugContext } from "./sampledebugContext";
import { Executor } from "../../utils/executor";
import { expect } from "chai";

class OutlookSignatureTestCase extends CaseFactory {
  override async onAfterCreate(
    sampledebugContext: SampledebugContext
  ): Promise<void> {
    const { success } = await Executor.execute(
      "npm install",
      sampledebugContext.projectPath,
      process.env,
      undefined,
      "npm warn"
    );
    expect(success).to.be.true;
    const childProcess = Executor.spawnCommand(
      sampledebugContext.projectPath,
      "npm",
      ["run", "start:xml"],
      (data) => {
        console.log(data.toString());
      },
      (error) => {
        console.error(error);
      }
    );
    await new Promise((resolve) => {
      setTimeout(() => {
        resolve(void 0);
      }, Timeout.installWait);
    });
    childProcess.kill();
  }
}

new OutlookSignatureTestCase(
  TemplateProject.OutlookSignature,
  21044484,
  24121523,
  "v-ivanchen@microsoft.com",
  [],
  {
    skipInit: true,
    repoPath: "./resource/Samples",
  }
).test();
