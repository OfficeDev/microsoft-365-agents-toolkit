// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { ok } from "@microsoft/teamsfx-api";
import axios from "axios";
import { assert } from "chai";
import fs from "fs-extra";
import { createSandbox } from "sinon";
import { setTools } from "../../../../src/common/globalVars";
import { ArmDeployDriver } from "../../../../src/component/driver/arm/deploy";
import { ArmDeployImpl } from "../../../../src/component/driver/arm/deployImpl";
import { azureClientHelper } from "../../../../src/component/utils/azureClient";
import { cpUtils } from "../../../../src/component/utils/depsChecker/cpUtils";
import {
  MockedAzureAccountProvider,
  MockedM365Provider,
  MockLogProvider,
  MockTelemetryReporter,
  MockTools,
  MockUserInteraction,
} from "../../../core/utils";

describe("Arm driver deploy", () => {
  const sandbox = createSandbox();
  const tools = new MockTools();
  setTools(tools);
  const mockedDriverContext: any = {
    m365TokenProvider: new MockedM365Provider(),
    azureAccountProvider: new MockedAzureAccountProvider(),
    telemetryReporter: new MockTelemetryReporter(),
    ui: new MockUserInteraction(),
    logProvider: new MockLogProvider(),
    projectPath: "./",
  };
  const driver = new ArmDeployDriver();

  const bicepCliVersion = "v0.9.1";
  beforeEach(() => {});

  afterEach(() => {
    sandbox.restore();
  });

  for (const actionMethod of ["run", "execute"]) {
    it(`happy path for ${actionMethod}`, async () => {
      sandbox.stub(fs, "readFile").resolves("{}" as any);
      sandbox.stub(cpUtils, "executeCommand").resolves("{}" as any);
      const deployRes = ok({
        mockKey: {
          type: "string",
          value: "mockValue",
        },
      });
      sandbox.stub(azureClientHelper, "createRmClient").resolves({} as any);
      sandbox.stub(ArmDeployImpl.prototype, "executeDeployment").resolves(deployRes as any);
      sandbox.stub(ArmDeployImpl.prototype, "ensureBicepCli").resolves("bicep");
      const fakeAxiosInstance = axios.create();
      sandbox.stub(axios, "create").returns(fakeAxiosInstance);
      sandbox.stub(fakeAxiosInstance, "get").resolves({
        status: 200,
        data: "",
      });
      let res: any;
      let deployArgs = {
        subscriptionId: "00000000-0000-0000-0000-000000000000",
        resourceGroupName: "mock-group",
        bicepCliVersion: bicepCliVersion,
        templates: [
          {
            path: "mock-template.bicep",
            parameters: "mock-parameters.json",
            deploymentName: "mock-deployment",
          },
          {
            path: "mock-template2.json",
            deploymentName: "mock-deployment2",
          },
          {
            path: "mock-template3.json",
            parameters: "mock-parameters3.json",
            deploymentName: "mock-deployment3",
          },
        ],
      };

      res = await driver.execute(deployArgs, mockedDriverContext);
      assert.isTrue(res.result.isOk());

      deployArgs = {
        subscriptionId: "00000000-0000-0000-0000-000000000000",
        resourceGroupName: "mock-group",
        bicepCliVersion: "",
        templates: [
          {
            path: "mock-template.json",
            parameters: "mock-parameters.json",
            deploymentName: "mock-deployment",
          },
        ],
      };
      if (actionMethod === "run") {
        res = await driver.execute(deployArgs, mockedDriverContext);
        assert.isTrue(res.result.isOk());
      } else {
        res = await driver.execute(deployArgs, mockedDriverContext);
        assert.isTrue(res.result.isOk());
      }
    });
  }

  it("invalid parameters", async () => {
    sandbox.stub(fs, "readFile").resolves("{}" as any);
    sandbox.stub(cpUtils, "executeCommand").resolves("{}" as any);
    let deployArgs = {
      subscriptionId: "",
      resourceGroupName: "",
      bicepCliVersion: "",
      templates: [
        {
          path: "mock-template",
          parameters: "mock-parameters",
          deploymentName: "",
        },
      ],
    } as any;
    let res = await driver.execute(deployArgs, mockedDriverContext);
    assert.isTrue(res.result.isErr());

    deployArgs = {
      subscriptionId: "00000000-0000-0000-0000-000000000000",
      resourceGroupName: "mock-group",
      bicepCliVersion: "",
      templates: [],
    } as any;
    res = await driver.execute(deployArgs, mockedDriverContext);
    assert.isTrue(res.result.isErr());

    deployArgs = {
      subscriptionId: "00000000-0000-0000-0000-000000000000",
      resourceGroupName: "mock-group",
      bicepCliVersion: "",
      templates: [
        {
          path: "C:/mock-template",
          parameters: "",
          deploymentName: "",
        },
      ],
    } as any;
    res = await driver.execute(deployArgs, mockedDriverContext);
    assert.isTrue(res.result.isErr());
  });

  it("deploy error", async () => {
    sandbox.stub(fs, "readFile").resolves("{}" as any);
    sandbox.stub(cpUtils, "executeCommand").resolves("{}" as any);
    sandbox
      .stub(ArmDeployImpl.prototype, "innerExecuteDeployment")
      .rejects(new Error("mocked deploy error"));
    sandbox.stub(ArmDeployImpl.prototype, "ensureBicepCli").resolves();
    const deployArgs = {
      subscriptionId: "00000000-0000-0000-0000-000000000000",
      resourceGroupName: "mock-group",
      bicepCliVersion: bicepCliVersion,
      templates: [
        {
          path: "mock-template.bicep",
          parameters: "mock-parameters.json",
          deploymentName: "mock-deployment",
        },
        {
          path: "mock-template2.json",
          parameters: "mock-parameters2.json",
          deploymentName: "mock-deployment2",
        },
      ],
    };

    const res = await driver.execute(deployArgs, mockedDriverContext);
    assert.isTrue(res.result.isErr());
  });

  it("error handle", async () => {
    sandbox.stub(ArmDeployImpl.prototype, "run").throws("mocked deploy error");

    const res = await driver.execute({} as any, mockedDriverContext);
    assert.isTrue(res.result.isErr());
  });
});
