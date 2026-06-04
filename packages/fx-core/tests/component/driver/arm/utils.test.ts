// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { ResourceManagementClient } from "@azure/arm-resources";
import { ok } from "@microsoft/teamsfx-api";
import { assert } from "chai";
import fs from "fs-extra";
import { createSandbox } from "sinon";
import { ConstantString } from "../../../../src/common/constants";
import { setTools } from "../../../../src/common/globalVars";
import { ArmDeployImpl } from "../../../../src/component/driver/arm/deployImpl";
import {
  ArmErrorHandle,
  armErrorHandleDeps,
} from "../../../../src/component/driver/arm/util/handleError";
import { convertOutputs, getFileExtension } from "../../../../src/component/driver/arm/util/util";
import { getAbsolutePath } from "../../../../src/component/utils/common";
import { cpUtils } from "../../../../src/component/utils/depsChecker/cpUtils";
import { MissingEnvironmentVariablesError } from "../../../../src/error";
import {
  CompileBicepError,
  DeployArmError,
  GetArmDeploymentError,
} from "../../../../src/error/arm";
import { ResourceGroupNotExistError } from "../../../../src/error/azure";
import {
  MockedAzureAccountProvider,
  MockedM365Provider,
  MockLogProvider,
  MockTelemetryReporter,
  MockTools,
  MockUserInteraction,
  MyTokenCredential,
} from "../../../core/utils";

describe("utils test", () => {
  const sandbox = createSandbox();
  const tools = new MockTools();
  setTools(tools);

  beforeEach(() => {});

  afterEach(() => {
    sandbox.restore();
  });

  it("getAbsolutePath empty", () => {
    const relativeOrAbsolutePath = undefined;
    const projectPath = undefined;
    const res = getAbsolutePath(
      relativeOrAbsolutePath as unknown as string,
      projectPath as unknown as string
    );
    assert.equal(res, ".");
  });

  it("getAbsolutePath empty", () => {
    const relativeOrAbsolutePath = undefined;
    const projectPath = undefined;
    const res = getAbsolutePath(
      relativeOrAbsolutePath as unknown as string,
      projectPath as unknown as string
    );
    assert.equal(res, ".");
  });

  it("getAbsolutePath absolute path", () => {
    const relativeOrAbsolutePath = "C:/a";
    const projectPath = "";
    const res = getAbsolutePath(relativeOrAbsolutePath, projectPath);
    assert.equal(relativeOrAbsolutePath, res);
  });

  it("getFileExtension empty", () => {
    const res = getFileExtension("");
    assert.isEmpty(res);
  });

  it("convert output", () => {
    const mockOutput = [
      {
        tabOutput: {
          type: "Object",
          value: {
            keyA: {
              type: "string",
              value: "valueA",
            },
            KeyB: 1,
          },
        },
      },
    ];
    const res = convertOutputs(mockOutput);
    assert.isNotEmpty(res);
  });
});

describe("arm deploy error handle test", () => {
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

  beforeEach(() => {});

  afterEach(() => {
    sandbox.restore();
  });

  it("handleArmDeploymentError case 1", async () => {
    let mockError = {
      code: "InvalidTemplateDeployment",
      message:
        "The template deployment 'Create-resources-for-tab' is not valid according to the validation procedure. The tracking id is '7da4fab7-ed36-4abc-9772-e2f90a0587a4'. See inner errors for details.",
      details: {
        error: {
          code: "ValidationForResourceFailed",
          message:
            "Validation failed for a resource. Check 'Error.Details[0]' for more information.",
          details: [
            {
              code: "MaxNumberOfServerFarmsInSkuPerSubscription",
              message: "The maximum number of Free ServerFarms allowed in a Subscription is 10.",
            },
          ],
        },
      },
    };

    let res = await ArmErrorHandle.handleArmDeploymentError(mockError, {
      ctx: { logProvider: new MockLogProvider() },
      deploymentName: "mockDeployName",
      resourceGroupName: "mockRG",
    } as any);
    assert.isTrue(res.isErr());

    mockError = {
      code: "InvalidTemplateDeployment",
      message:
        "The template deployment 'Create-resources-for-tab' is not valid according to the validation procedure. The tracking id is '7da4fab7-ed36-4abc-9772-e2f90a0587a4'. See inner errors for details.",
      details: {
        code: "ValidationForResourceFailed",
        message: "Validation failed for a resource. Check 'Error.Details[0]' for more information.",
        details: [
          {
            code: "MaxNumberOfServerFarmsInSkuPerSubscription",
            message: "The maximum number of Free ServerFarms allowed in a Subscription is 10.",
          },
        ],
      },
    } as any;
    res = await ArmErrorHandle.handleArmDeploymentError(mockError, {
      ctx: { logProvider: new MockLogProvider() },
      deploymentName: "mockDeployName",
      resourceGroupName: "mockRG",
    } as any);
    assert.isTrue(res.isErr());
    if (res.isErr()) {
      assert.isTrue(res.error instanceof DeployArmError);
    }

    mockError = {
      code: "ResourceGroupNotFound",
      message:
        "The template deployment 'Create-resources-for-tab' is not valid according to the validation procedure. The tracking id is '7da4fab7-ed36-4abc-9772-e2f90a0587a4'. See inner errors for details.",
      details: {
        code: "ValidationForResourceFailed",
        message: "Validation failed for a resource. Check 'Error.Details[0]' for more information.",
        details: [
          {
            code: "MaxNumberOfServerFarmsInSkuPerSubscription",
            message: "The maximum number of Free ServerFarms allowed in a Subscription is 10.",
          },
        ],
      },
    } as any;
    res = await ArmErrorHandle.handleArmDeploymentError(mockError, {
      ctx: { logProvider: new MockLogProvider() },
      deploymentName: "mockDeployName",
      resourceGroupName: "mockRG",
      client: { subscriptionId: "mockSubId" },
    } as any);
    assert.isTrue(res.isErr());
    if (res.isErr()) {
      assert.isTrue(res.error instanceof ResourceGroupNotExistError);
    }
  });

  it("handleArmDeploymentError case 2: no deploymentError", async () => {
    const client = new ResourceManagementClient(new MyTokenCredential(), "id");
    sandbox.stub(client.deployments, "get").resolves({});
    const mockError = {
      code: "OtherCode",
      message:
        "The template deployment 'Create-resources-for-tab' is not valid according to the validation procedure. The tracking id is '7da4fab7-ed36-4abc-9772-e2f90a0587a4'. See inner errors for details.",
    };
    const res = await ArmErrorHandle.handleArmDeploymentError(mockError, {
      ctx: { logProvider: new MockLogProvider() },
      deploymentName: "mockDeployName",
      resourceGroupName: "mockRG",
      client: client,
    } as any);
    assert.isTrue(res.isErr());
    if (res.isErr()) {
      assert.isTrue(res.error instanceof DeployArmError);
    }
  });

  it("handleArmDeploymentError case 3: getDeploymentError without subErrors", async () => {
    const client = new ResourceManagementClient(new MyTokenCredential(), "id");
    sandbox.stub(ArmErrorHandle, "wrapGetDeploymentError").resolves(
      ok({
        error: {
          code: "MockError",
          message: "MockErrorMessage",
        },
      })
    );
    const mockError = {
      code: "RawMockError",
      message: "RawMockErrorMessasge",
    };
    const res = await ArmErrorHandle.handleArmDeploymentError(mockError, {
      ctx: { logProvider: new MockLogProvider() },
      deploymentName: "mockDeployName",
      resourceGroupName: "mockRG",
      client: client,
    } as any);
    assert.isTrue(res.isErr());
    if (res.isErr()) {
      assert.isTrue(res.error instanceof DeployArmError);
    }
  });

  it("handleArmDeploymentError case 4: getDeploymentError with subErrors", async () => {
    const client = new ResourceManagementClient(new MyTokenCredential(), "id");
    sandbox
      .stub(ArmErrorHandle, "wrapGetDeploymentError")
      .resolves(ok({ subErrors: { module1: "value1" } }));
    const mockError = {
      code: "RawMockError",
      message: "RawMockErrorMessasge",
    };
    const res = await ArmErrorHandle.handleArmDeploymentError(mockError, {
      ctx: { logProvider: new MockLogProvider() },
      deploymentName: "mockDeployName",
      resourceGroupName: "mockRG",
      client: client,
    } as any);
    assert.isTrue(res.isErr());
    if (res.isErr()) {
      assert.isTrue(res.error instanceof DeployArmError);
    }
  });

  it("handleArmDeploymentError case 5: getDeploymentError throws error", async () => {
    const client = new ResourceManagementClient(new MyTokenCredential(), "id");
    sandbox
      .stub(ArmErrorHandle, "getDeploymentError")
      .throws({ code: "GetDeploymentError", message: "GetDeploymentErrorMessage" });
    const mockError = {
      code: "RawMockError",
      message: "RawMockErrorMessasge",
    };
    const res = await ArmErrorHandle.handleArmDeploymentError(mockError, {
      ctx: { logProvider: new MockLogProvider() },
      deploymentName: "mockDeployName",
      resourceGroupName: "mockRG",
      client: client,
    } as any);
    assert.isTrue(res.isErr());
    if (res.isErr()) {
      assert.isTrue(res.error instanceof GetArmDeploymentError);
    }
  });

  it("deployTemplate throw DeployTemplate FxError", async () => {
    const deployArgs = {
      subscriptionId: "00000000-0000-0000-0000-000000000000",
      resourceGroupName: "mock-group",
      bicepCliVersion: "",
      templates: [],
    } as any;
    const impl = new ArmDeployImpl(deployArgs, mockedDriverContext);
    sandbox
      .stub(impl, "getDeployTemplate")
      .throws(new CompileBicepError(".", new Error("compile error")));
    mockedDriverContext.createProgressBar = () => {};
    const res = await impl.deployTemplate({
      path: "",
      parameters: "",
      deploymentName: "mkdpn",
    });
    assert.isTrue(res.isErr());
    if (res.isErr()) {
      assert.isTrue(res.error instanceof CompileBicepError);
    }
  });

  it("deployTemplate throw unresolved env variable FxError", async () => {
    const deployArgs = {
      subscriptionId: "00000000-0000-0000-0000-000000000000",
      resourceGroupName: "mock-group",
      bicepCliVersion: "",
      templates: [],
    } as any;
    const impl = new ArmDeployImpl(deployArgs, mockedDriverContext);
    const parameterContents = `
    {
      "$schema": "https://schema.management.azure.com/schemas/2015-01-01/deploymentParameters.json#",
      "contentVersion": "1.0.0.0",
      "parameters": {
        "resourceBaseName": {
          "value": "bot\${{RESOURCE_SUFFIX}}"
        },
        "botAadAppClientId": {
          "value": "\${{BOT_ID}}"
        },
        "botAadAppClientSecret": {
          "value": "\${{SECRET_BOT_PASSWORD}}"
        },
        "webAppSKU": {
          "value": "B1"
        },
      }
    }
    `;
    sandbox.stub(fs, "readFile").resolves(parameterContents as any);
    mockedDriverContext.createProgressBar = () => {};
    const res = await impl.deployTemplate({
      path: "",
      parameters: "mockParam",
      deploymentName: "mkdpn",
    });
    assert.isTrue(res.isErr());
    if (res.isErr()) {
      assert.isTrue(res.error instanceof MissingEnvironmentVariablesError);
      assert.isTrue(res.error.message.includes("RESOURCE_SUFFIX"));
      assert.isTrue(res.error.message.includes("BOT_ID"));
      assert.isTrue(res.error.message.includes("SECRET_BOT_PASSWORD"));
    }
  });

  it("compileBicepToJson throw Error", async () => {
    const deployArgs = {
      subscriptionId: "00000000-0000-0000-0000-000000000000",
      resourceGroupName: "mock-group",
      bicepCliVersion: "",
      templates: [],
    } as any;
    mockedDriverContext.createProgressBar = () => {};
    sandbox.stub(cpUtils, "executeCommand").throws(new Error("compile error"));
    const impl = new ArmDeployImpl(deployArgs, mockedDriverContext);
    try {
      await impl.compileBicepToJson("");
      assert.fail("should not reach here");
    } catch (e) {
      assert.isTrue(e instanceof CompileBicepError);
    }
  });
});

describe("getDeploymentError", () => {
  const sandbox = createSandbox();
  const tools = new MockTools();
  setTools(tools);
  const deployCtx = {
    ctx: { logProvider: new MockLogProvider() },
    deploymentName: "mockDeployName",
    resourceGroupName: "mockRG",
    deploymentStartTime: Date.now(),
  } as any;
  beforeEach(() => {});

  afterEach(() => {
    sandbox.restore();
  });

  it("throw error", async () => {
    sandbox
      .stub(armErrorHandleDeps, "innerGetDeploymentError")
      .throws({ code: ConstantString.DeploymentNotFound });
    try {
      await ArmErrorHandle.getDeploymentError(
        deployCtx,
        deployCtx.resourceGroupName,
        deployCtx.deploymentName
      );
      assert.fail("should not reach here");
    } catch (error) {
      assert.isTrue(error.code === ConstantString.DeploymentNotFound);
    }
  });

  it("get error:empty", async () => {
    sandbox
      .stub(armErrorHandleDeps, "innerGetDeploymentError")
      .throws({ code: ConstantString.DeploymentNotFound });
    const res = await ArmErrorHandle.getDeploymentError(
      deployCtx,
      deployCtx.resourceGroupName,
      "mockDeploymentName"
    );
    assert.isUndefined(res);
  });

  it("timestamp is less than startTime", async () => {
    sandbox.stub(armErrorHandleDeps, "innerGetDeploymentError").resolves({
      properties: {
        timestamp: new Date(deployCtx.deploymentStartTime - 1000),
      },
    } as any);
    const res = await ArmErrorHandle.getDeploymentError(
      deployCtx,
      deployCtx.resourceGroupName,
      deployCtx.deploymentName
    );
    assert.isUndefined(res);
  });

  it("error is empty", async () => {
    sandbox.stub(armErrorHandleDeps, "innerGetDeploymentError").resolves({
      properties: {
        timestamp: new Date(),
      },
    } as any);
    const res = await ArmErrorHandle.getDeploymentError(
      deployCtx,
      deployCtx.resourceGroupName,
      deployCtx.deploymentName
    );
    assert.isUndefined(res);
  });

  it("error not empty", async () => {
    sandbox.stub(armErrorHandleDeps, "innerGetDeploymentError").resolves({
      properties: {
        error: {
          message: "mockMessage",
        },
      },
    } as any);
    sandbox.stub(armErrorHandleDeps, "innerGetDeploymentOperations").resolves([
      {
        properties: {
          targetResource: {
            resourceName: "mockResourceName",
          },
          statusMessage: {
            error: {},
          },
        },
      },
    ] as any);
    const res = await ArmErrorHandle.getDeploymentError(
      deployCtx,
      deployCtx.resourceGroupName,
      deployCtx.deploymentName
    );
    assert.isNotEmpty(res);
  });

  it("error not empty and nested error", async () => {
    // sandbox.stub(innerHandleError, "innerGetDeploymentError").onFirstCall
    sandbox
      .stub(armErrorHandleDeps, "innerGetDeploymentError")
      .onFirstCall()
      .resolves({
        properties: {
          error: {
            message: "mockMessage",
          },
        },
      } as any)
      .onSecondCall()
      .throws({ code: ConstantString.DeploymentNotFound });
    sandbox.stub(armErrorHandleDeps, "innerGetDeploymentOperations").resolves([
      {
        id: "mockId",
        properties: {
          targetResource: {
            resourceType: ConstantString.DeploymentResourceType,
            resourceName: "mockResourceName2",
            id: "/resourceGroups/mockGroup2/mockId2",
          },
          statusMessage: {
            error: {},
          },
        },
      },
    ] as any);
    const res = await ArmErrorHandle.getDeploymentError(
      deployCtx,
      deployCtx.resourceGroupName,
      deployCtx.deploymentName
    );
    assert.isNotEmpty(res);
  });
});

describe("formattedDeploymentError Status", () => {
  const mocker = createSandbox();

  beforeEach(async () => {});

  afterEach(async () => {
    mocker.restore();
  });

  it("formattedDeploymentError OK", async () => {
    const errors = {
      error: {
        code: "OutsideError",
        message: "out side error",
      },
      subErrors: {
        botProvision: {
          error: {
            code: "BotError",
            message: "bot error",
          },
          inner: {
            error: {
              code: "BotInnerError",
              message: "bot inner error",
            },
            subErrors: {
              skuError: {
                error: {
                  code: "MaxNumberOfServerFarmsInSkuPerSubscription",
                  message: "The maximum number of Free ServerFarms allowed in a Subscription is 10",
                },
              },
              evaluationError: {
                error: {
                  code: "DeploymentOperationFailed",
                  message:
                    "Template output evaluation skipped: at least one resource deployment operation failed. Please list deployment operations for details. Please see https://aka.ms/DeployOperations for usage details.",
                },
              },
            },
          },
        },
      },
    };
    const res = ArmErrorHandle.formattedDeploymentError(errors);
    assert.deepEqual(res, {
      botProvision: {
        skuError: {
          code: "MaxNumberOfServerFarmsInSkuPerSubscription",
          message: "The maximum number of Free ServerFarms allowed in a Subscription is 10",
        },
      },
    });
  });
});
