// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import "mocha";
import * as sinon from "sinon";
import mockedEnv, { RestoreFn } from "mocked-env";
import * as chai from "chai";
import chaiAsPromised from "chai-as-promised";
import { CreateSharePointEmbeddedContainerTypeDriver } from "../../../../src/component/driver/sharePointEmbeddedContainerType/create";
import { SPEContainerTypeAppClient } from "../../../../src/component/driver/sharePointEmbeddedContainerType/utility/speContainerTypeAppClient";
import {
  SPContainerBillingStatus,
  SPContainerTypeSettingsOverride,
  ISharePointEmbeddedContainerType,
  SharePointEmbeddedContainerTypeSettings,
  SPContainerTypeBillingClassification,
} from "../../../../src/component/driver/sharePointEmbeddedContainerType/interface/sharePointEmbeddedContainerType";
import { MockedM365Provider } from "../../../core/utils";
import {
  MockedLogProvider,
  MockedTelemetryReporter,
  MockedUserInteraction,
} from "../../../plugins/solution/util";
import { HttpServerError, InvalidActionInputError } from "../../../../src/error";
import { OutputEnvironmentVariableUndefinedError } from "../../../../src/component/driver/error/outputEnvironmentVariableUndefinedError";
chai.use(chaiAsPromised);
const expect = chai.expect;

const outputKeys = {
  containerTypeId: "ContainerTypeId",
};

const outputEnvVarNames = new Map<string, string>(Object.entries(outputKeys));

describe("sharePointEmbeddedContainerTypeCreate", async () => {
  const speContainerTypeDriver = new CreateSharePointEmbeddedContainerTypeDriver();

  const mockedDriverContext: any = {
    m365TokenProvider: new MockedM365Provider(),
    ui: new MockedUserInteraction(),
    projectPath: "test",
  };

  const expectedContainerTypeId = "247e5bd2-edc2-4ee7-a9a3-0064eaf77ccb";
  const containerTypeName = "TestContainerType";
  const containerTypeOwningAppId = "00000000-0000-0000-0000-000000000001";
  const containerTypeBillingClassification: SPContainerTypeBillingClassification =
    SPContainerTypeBillingClassification.standard;
  const containerTypeBillingStatus: SPContainerBillingStatus = SPContainerBillingStatus.valid;
  const containerTypeCreatedDateTime: Date = new Date("2024-06-24T12:00:00Z");
  const containerTypeExpirationDateTime: Date = new Date("2025-06-24T12:00:00Z");
  const containerTypeSettings: SharePointEmbeddedContainerTypeSettings = {
    sharingCapability: "external",
    urlTemplate: "https://example.com/{id}",
    isDiscoverabilityEnabled: true,
    isSearchEnabled: true,
    isItemVersioningEnabled: true,
    itemMajorVersionLimit: 10,
    maxStoragePerContainerInBytes: 1000000000,
    isSharingRestricted: false,
    consumingTenantOverridables: SPContainerTypeSettingsOverride.none,
  }; // settings for the container type
  const containerTypeEtag = "eTag"; // etag for optimistic concurrency control

  const testContainerType: ISharePointEmbeddedContainerType = {
    id: expectedContainerTypeId,
    name: containerTypeName,
    owningAppId: containerTypeOwningAppId,
    billingClassification: containerTypeBillingClassification,
    billingStatus: containerTypeBillingStatus,
    createdDateTime: containerTypeCreatedDateTime,
    expirationDateTime: containerTypeExpirationDateTime,
    settings: containerTypeSettings,
    etag: containerTypeEtag,
  };

  let envRestore: RestoreFn | undefined;

  afterEach(() => {
    sinon.restore();
    if (envRestore) {
      envRestore();
      envRestore = undefined;
    }
  });

  it("should return error if arguments are invalid", async () => {
    // name is invalid
    let args: any = {
      name: 123,
      owningApplicationId: "app-id",
      billingClassification: SPContainerTypeBillingClassification.standard,
      discoverable: true,
    };
    let result = await speContainerTypeDriver.execute(args, mockedDriverContext, outputEnvVarNames);
    expect(result.result.isErr()).to.be.true;
    expect(result.result._unsafeUnwrapErr()).to.be.instanceOf(InvalidActionInputError);

    // owningApplicationId is invalid
    args = {
      name: "test",
      billingClassification: SPContainerTypeBillingClassification.standard,
      discoverable: true,
      owningApplicationId: 123,
    };
    result = await speContainerTypeDriver.execute(args, mockedDriverContext, outputEnvVarNames);
    expect(result.result.isErr()).to.be.true;
    expect(result.result._unsafeUnwrapErr()).to.be.instanceOf(InvalidActionInputError);

    // billingClassification invalid
    args = {
      name: "test",
      owningApplicationId: "app-id",
      billingClassification: "invalid",
      discoverable: true,
    };
    result = await speContainerTypeDriver.execute(args, mockedDriverContext, outputEnvVarNames);
    expect(result.result.isErr()).to.be.true;
    expect(result.result._unsafeUnwrapErr()).to.be.instanceOf(InvalidActionInputError);

    // discoverable invalid
    args = {
      name: "test",
      owningApplicationId: "app-id",
      billingClassification: SPContainerTypeBillingClassification.standard,
      discoverable: "notBoolean",
    };
    result = await speContainerTypeDriver.execute(args, mockedDriverContext, outputEnvVarNames);
    expect(result.result.isErr()).to.be.true;
    expect(result.result._unsafeUnwrapErr()).to.be.instanceOf(InvalidActionInputError);
  });

  it("should return error if arguments are missing", async () => {
    // name missing
    let args: any = {
      owningApplicationId: "app-id",
      billingClassification: SPContainerTypeBillingClassification.standard,
      discoverable: true,
    };
    let result = await speContainerTypeDriver.execute(args, mockedDriverContext, outputEnvVarNames);
    expect(result.result.isErr()).to.be.true;
    expect(result.result._unsafeUnwrapErr()).to.be.instanceOf(InvalidActionInputError);

    // owningApplicationId missing
    args = {
      name: "test",
      billingClassification: SPContainerTypeBillingClassification.standard,
      discoverable: true,
    };
    result = await speContainerTypeDriver.execute(args, mockedDriverContext, outputEnvVarNames);
    expect(result.result.isErr()).to.be.true;
    expect(result.result._unsafeUnwrapErr()).to.be.instanceOf(InvalidActionInputError);

    // billingClassification missing
    args = {
      name: "test",
      owningApplicationId: "app-id",
      discoverable: true,
    };
    result = await speContainerTypeDriver.execute(args, mockedDriverContext, outputEnvVarNames);
    expect(result.result.isErr()).to.be.true;
    expect(result.result._unsafeUnwrapErr()).to.be.instanceOf(InvalidActionInputError);

    // discoverable missing
    args = {
      name: "test",
      owningApplicationId: "app-id",
      billingClassification: SPContainerTypeBillingClassification.standard,
    };
    result = await speContainerTypeDriver.execute(args, mockedDriverContext, outputEnvVarNames);
    expect(result.result.isErr()).to.be.true;
    expect(result.result._unsafeUnwrapErr()).to.be.instanceOf(InvalidActionInputError);
  });

  it("should return error if outputEnvVarNames is undefined", async () => {
    const args = {
      name: "test",
      owningApplicationId: "app-id",
      billingClassification: SPContainerTypeBillingClassification.standard,
      discoverable: true,
    };
    const result = await speContainerTypeDriver.execute(args, mockedDriverContext);
    expect(result.result.isErr()).to.be.true;
    expect(result.result._unsafeUnwrapErr()).to.be.instanceOf(
      OutputEnvironmentVariableUndefinedError
    );
  });

  it("should create new SPEContainerType with empty .env", async () => {
    sinon
      .stub(SPEContainerTypeAppClient.prototype, "createSPEContainerType")
      .resolves(testContainerType as ISharePointEmbeddedContainerType);

    const args: any = {
      name: containerTypeName,
      owningApplicationId: containerTypeOwningAppId,
      billingClassification: containerTypeBillingClassification,
      discoverable: true,
    };

    const result = await speContainerTypeDriver.execute(
      args,
      mockedDriverContext,
      outputEnvVarNames
    );

    expect(result.result.isOk()).to.be.true;
    expect(result.result._unsafeUnwrap().get(outputKeys.containerTypeId)).to.equal(
      expectedContainerTypeId
    );
  });

  it("should output to specific environment variable based on writeToEnvironmentFile declaration", async () => {
    sinon
      .stub(SPEContainerTypeAppClient.prototype, "createSPEContainerType")
      .resolves(testContainerType as ISharePointEmbeddedContainerType);

    const args: any = {
      name: containerTypeName,
      owningApplicationId: containerTypeOwningAppId,
      billingClassification: containerTypeBillingClassification,
      discoverable: true,
    };

    const outputEnvVarNames = new Map<string, string>(
      Object.entries({
        containerTypeId: "MY_CONTAINER_TYPE_ID",
      })
    );

    const result = await speContainerTypeDriver.execute(
      args,
      mockedDriverContext,
      outputEnvVarNames
    );

    expect(result.result.isOk()).to.be.true;
    expect(result.result._unsafeUnwrap().get("MY_CONTAINER_TYPE_ID")).to.equal(
      expectedContainerTypeId
    );
  });

  it("should use existing SPE Container Type when CONTAINER_TYPE_ID exists", async () => {
    sinon
      .stub(SPEContainerTypeAppClient.prototype, "createSPEContainerType")
      .rejects("createSPEContainerType should not be called");

    envRestore = mockedEnv({
      [outputKeys.containerTypeId]: "5eb48390-7c1a-48af-be78-fe35cc24e956",
    });

    const args: any = {
      name: containerTypeName,
      owningApplicationId: containerTypeOwningAppId,
      billingClassification: containerTypeBillingClassification,
      discoverable: true,
    };

    const result = await speContainerTypeDriver.execute(
      args,
      mockedDriverContext,
      outputEnvVarNames
    );

    expect(result.result.isOk()).to.be.true;
    expect(result.result._unsafeUnwrap().size).to.equal(0);
    expect(result.summaries.length).to.equal(0);
  });

  it("should throw error when SPEContainerTypeClient fails", async () => {
    sinon.stub(SPEContainerTypeAppClient.prototype, "createSPEContainerType").rejects({
      isAxiosError: true,
      response: {
        status: 500,
        data: {
          error: {
            code: "InternalServerError",
            message: "Internal server error",
          },
        },
      },
    });

    const args: any = {
      name: containerTypeName,
      owningApplicationId: containerTypeOwningAppId,
      billingClassification: containerTypeBillingClassification,
      discoverable: true,
    };

    const result = await speContainerTypeDriver.execute(
      args,
      mockedDriverContext,
      outputEnvVarNames
    );

    expect(result.result.isErr()).to.be.true;
    expect(result.result._unsafeUnwrapErr())
      .is.instanceOf(HttpServerError)
      .and.has.property("message")
      .and.equals(
        'A http server error occurred while performing the sharePointEmbeddedContainerType/create task. Try again later. The error response is: {"error":{"code":"InternalServerError","message":"Internal server error"}}'
      );
  });
});
