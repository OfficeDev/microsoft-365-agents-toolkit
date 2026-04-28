import { ResourceManagementClient } from "@azure/arm-resources";
import { ok, Platform } from "@microsoft/teamsfx-api";
import { assert } from "chai";
import "mocha";
import * as sinon from "sinon";
import { resourceGroupHelper } from "../../src/component/utils/ResourceGroupHelper";
import { setTools, TOOLS } from "../../src/common/globalVars";
import { MockTools, MyTokenCredential } from "../core/utils";
import * as armResources from "@azure/arm-resources";
import * as armSubscriptions from "@azure/arm-subscriptions";
import {
  CheckResourceGroupExistenceError,
  CreateResourceGroupError,
  GetResourceGroupError,
  ListResourceGroupLocationsError,
  ListResourceGroupsError,
} from "../../src/error/azure";
import { SubscriptionClient } from "@azure/arm-subscriptions";

describe("resouce group helper test", () => {
  const sandbox = sinon.createSandbox();
  const tools = new MockTools();
  setTools(tools);
  beforeEach(() => {
    sandbox.stub(tools.tokenProvider.azureAccountProvider, "setSubscription").resolves();
  });
  afterEach(() => {
    sandbox.restore();
  });
  it("askResourceGroupInfoV3", async () => {
    sandbox.stub(resourceGroupHelper, "listResourceGroups").resolves(ok([["rg1", "loc1"]]));
    sandbox.stub(resourceGroupHelper, "getLocations").resolves(ok(["loc1"]));
    sandbox
      .stub(TOOLS.ui, "selectOption")
      .resolves(ok({ type: "success", result: { id: "rg1", label: "loc1" } }));
    const mockResourceManagementClient = new ResourceManagementClient(
      new MyTokenCredential(),
      "id"
    );
    const res = await resourceGroupHelper.askResourceGroupInfoV3(
      { platform: Platform.VSCode, projectPath: "" },
      tools.tokenProvider.azureAccountProvider,
      mockResourceManagementClient,
      "rg1"
    );
    if (res.isErr()) {
      console.error(res.error);
    }
    assert.isTrue(res.isOk());
  });

  it("createResourceGroup return undefined", async () => {
    const mockResourceManagementClient = new ResourceManagementClient(
      new MyTokenCredential(),
      "id"
    );
    sandbox.stub(armResources, "ResourceManagementClient").returns(mockResourceManagementClient);
    sandbox
      .stub(tools.tokenProvider.azureAccountProvider, "getIdentityCredentialAsync")
      .resolves(new MyTokenCredential());
    sandbox.stub(resourceGroupHelper, "checkResourceGroupExistence").resolves(ok(false));
    sandbox
      .stub(mockResourceManagementClient.resourceGroups, "createOrUpdate")
      .resolves({ name: undefined, location: "east us" });
    const res = await resourceGroupHelper.createNewResourceGroup(
      "mockRG",
      tools.tokenProvider.azureAccountProvider,
      "mockSubId",
      "east us"
    );
    assert.isTrue(res.isErr());
    if (res.isErr()) {
      assert.isTrue(res.error instanceof CreateResourceGroupError);
    }
  });

  it("createResourceGroup throw error", async () => {
    const mockResourceManagementClient = new ResourceManagementClient(
      new MyTokenCredential(),
      "id"
    );
    sandbox.stub(armResources, "ResourceManagementClient").returns(mockResourceManagementClient);
    sandbox
      .stub(tools.tokenProvider.azureAccountProvider, "getIdentityCredentialAsync")
      .resolves(new MyTokenCredential());
    sandbox.stub(resourceGroupHelper, "checkResourceGroupExistence").resolves(ok(false));
    sandbox
      .stub(mockResourceManagementClient.resourceGroups, "createOrUpdate")
      .rejects(new Error("test"));
    const res = await resourceGroupHelper.createNewResourceGroup(
      "mockRG",
      tools.tokenProvider.azureAccountProvider,
      "mockSubId",
      "east us"
    );
    assert.isTrue(res.isErr());
    if (res.isErr()) {
      assert.isTrue(res.error instanceof CreateResourceGroupError);
    }
  });

  it("createResourceGroup success", async () => {
    const mockResourceManagementClient = new ResourceManagementClient(
      new MyTokenCredential(),
      "id"
    );
    sandbox.stub(armResources, "ResourceManagementClient").returns(mockResourceManagementClient);
    sandbox
      .stub(tools.tokenProvider.azureAccountProvider, "getIdentityCredentialAsync")
      .resolves(new MyTokenCredential());
    sandbox.stub(resourceGroupHelper, "checkResourceGroupExistence").resolves(ok(false));
    sandbox
      .stub(mockResourceManagementClient.resourceGroups, "createOrUpdate")
      .resolves({ name: "mockRg", location: "east us" });
    const res = await resourceGroupHelper.createNewResourceGroup(
      "mockRG",
      tools.tokenProvider.azureAccountProvider,
      "mockSubId",
      "east us"
    );
    assert.isTrue(res.isOk());
  });

  it("checkResourceGroupExistence success", async () => {
    const mockResourceManagementClient = new ResourceManagementClient(
      new MyTokenCredential(),
      "id"
    );
    sandbox
      .stub(mockResourceManagementClient.resourceGroups, "checkExistence")
      .resolves({ body: true });
    const res = await resourceGroupHelper.checkResourceGroupExistence(
      "mockRG",
      mockResourceManagementClient
    );
    assert.isTrue(res.isOk());
    if (res.isOk()) {
      assert.isTrue(res.value);
    }
  });

  it("checkResourceGroupExistence throw Error", async () => {
    const mockResourceManagementClient = new ResourceManagementClient(
      new MyTokenCredential(),
      "id"
    );
    sandbox
      .stub(mockResourceManagementClient.resourceGroups, "checkExistence")
      .rejects(new Error("test"));
    const res = await resourceGroupHelper.checkResourceGroupExistence(
      "mockRG",
      mockResourceManagementClient
    );
    assert.isTrue(res.isErr());
    if (res.isErr()) {
      assert.isTrue(res.error instanceof CheckResourceGroupExistenceError);
    }
  });

  it("getResourceGroupInfo success", async () => {
    const mockResourceManagementClient = new ResourceManagementClient(
      new MyTokenCredential(),
      "id"
    );
    sandbox
      .stub(mockResourceManagementClient.resourceGroups, "get")
      .resolves({ name: "mockRG", location: "XXX" });
    const res = await resourceGroupHelper.getResourceGroupInfo(
      "mockRG",
      mockResourceManagementClient
    );
    assert.isTrue(res.isOk());
    if (res.isOk()) {
      assert.isTrue(res.value?.name === "mockRG");
    }
  });

  it("getResourceGroupInfo success return undefined", async () => {
    const mockResourceManagementClient = new ResourceManagementClient(
      new MyTokenCredential(),
      "id"
    );
    sandbox.stub(mockResourceManagementClient.resourceGroups, "get").resolves({ location: "XXX" });
    const res = await resourceGroupHelper.getResourceGroupInfo(
      "mockRG",
      mockResourceManagementClient
    );
    assert.isTrue(res.isOk());
    if (res.isOk()) {
      assert.isUndefined(res.value);
    }
  });

  it("getResourceGroupInfo throw Error", async () => {
    const mockResourceManagementClient = new ResourceManagementClient(
      new MyTokenCredential(),
      "id"
    );
    sandbox.stub(mockResourceManagementClient.resourceGroups, "get").rejects(new Error(""));
    const res = await resourceGroupHelper.getResourceGroupInfo(
      "mockRG",
      mockResourceManagementClient
    );
    assert.isTrue(res.isErr());
    if (res.isErr()) {
      assert.isTrue(res.error instanceof GetResourceGroupError);
    }
  });

  it("listResourceGroups success", async () => {
    const client = new ResourceManagementClient(new MyTokenCredential(), "id");
    const iterator = {
      next: sandbox
        .stub()
        .onFirstCall()
        .resolves({
          value: { name: "rg1", location: "east us" },
          done: false,
        })
        .onSecondCall()
        .resolves({
          value: { name: "rg2", location: "east us" },
          done: true,
        }),
      byPage: sandbox.stub().resolves([[{ name: "rg", location: "east us" }]]),
      [Symbol.asyncIterator]() {
        return this;
      },
    };
    sandbox.stub(client.resourceGroups, "list").returns(iterator);
    const res = await resourceGroupHelper.listResourceGroups(client);
    assert.isTrue(res.isOk());
    if (res.isOk()) {
      assert.isTrue(res.value.length === 2);
    }
  });

  it("listResourceGroups throw Error", async () => {
    const client = new ResourceManagementClient(new MyTokenCredential(), "id");
    const iterator = {
      next: sandbox.stub().rejects(new Error("test")),
      byPage: sandbox.stub().resolves([[{ name: "rg", location: "east us" }]]),
      [Symbol.asyncIterator]() {
        return this;
      },
    };
    sandbox.stub(client.resourceGroups, "list").returns(iterator);
    const res = await resourceGroupHelper.listResourceGroups(client);
    assert.isTrue(res.isErr());
    if (res.isErr()) {
      assert.isTrue(res.error instanceof ListResourceGroupsError);
    }
  });

  it("getLocations success", async () => {
    const rmClient = new ResourceManagementClient(new MyTokenCredential(), "id");
    const iterator = {
      next: sandbox
        .stub()
        .onFirstCall()
        .resolves({
          value: { displayName: "east us" },
          done: false,
        })
        .onSecondCall()
        .resolves({
          value: { displayName: "central us" },
          done: true,
        }),
      byPage: sandbox.stub().resolves([[{ displayName: "east us" }]]),
      [Symbol.asyncIterator]() {
        return this;
      },
    };
    const originalDescriptor = Object.getOwnPropertyDescriptor(
      armSubscriptions,
      "SubscriptionClient"
    )!;
    Object.defineProperty(armSubscriptions, "SubscriptionClient", {
      configurable: true,
      writable: true,
      value: function () {
        return { subscriptions: { listLocations: () => iterator } };
      },
    });
    sandbox.stub(tools.tokenProvider.azureAccountProvider, "getSelectedSubscription").resolves({
      subscriptionId: "mockSubId",
      subscriptionName: "mockSubName",
      tenantId: "mockTID",
    });
    sandbox.stub(rmClient.providers, "get").resolves({
      resourceTypes: [{ resourceType: "resourceGroups", locations: ["east us"] }],
    });
    sandbox
      .stub(tools.tokenProvider.azureAccountProvider, "getIdentityCredentialAsync")
      .resolves(new MyTokenCredential());
    try {
      const res = await resourceGroupHelper.getLocations(
        tools.tokenProvider.azureAccountProvider,
        rmClient
      );
      assert.isTrue(res.isOk());
      if (res.isOk()) {
        assert.isTrue(res.value.length === 1);
      }
    } finally {
      Object.defineProperty(armSubscriptions, "SubscriptionClient", originalDescriptor);
    }
  });

  it("getLocations return zero results", async () => {
    const subClient = new SubscriptionClient(new MyTokenCredential());
    const rmClient = new ResourceManagementClient(new MyTokenCredential(), "id");
    sandbox.stub(armSubscriptions, "SubscriptionClient").returns(subClient);
    const iterator = {
      next: sandbox
        .stub()
        .onFirstCall()
        .resolves({
          value: { displayName: "east us" },
          done: false,
        })
        .onSecondCall()
        .resolves({
          value: { displayName: "central us" },
          done: true,
        }),
      byPage: sandbox.stub().resolves([[{ displayName: "east us" }]]),
      [Symbol.asyncIterator]() {
        return this;
      },
    };
    sandbox.stub(subClient.subscriptions, "listLocations").returns(iterator);
    sandbox.stub(tools.tokenProvider.azureAccountProvider, "getSelectedSubscription").resolves({
      subscriptionId: "mockSubId",
      subscriptionName: "mockSubName",
      tenantId: "mockTID",
    });
    sandbox.stub(rmClient.providers, "get").resolves({
      resourceTypes: [],
    });
    sandbox
      .stub(tools.tokenProvider.azureAccountProvider, "getIdentityCredentialAsync")
      .resolves(new MyTokenCredential());
    const res = await resourceGroupHelper.getLocations(
      tools.tokenProvider.azureAccountProvider,
      rmClient
    );
    assert.isTrue(res.isErr());
    if (res.isErr()) {
      assert.isTrue(res.error instanceof ListResourceGroupLocationsError);
    }
  });

  it("getLocations throw Error", async () => {
    const subClient = new SubscriptionClient(new MyTokenCredential());
    const rmClient = new ResourceManagementClient(new MyTokenCredential(), "id");
    sandbox.stub(armSubscriptions, "SubscriptionClient").returns(subClient);
    const iterator = {
      next: sandbox.stub().rejects({ message: "test error" }),
      byPage: sandbox.stub().resolves([[{ displayName: "east us" }]]),
      [Symbol.asyncIterator]() {
        return this;
      },
    };
    sandbox.stub(subClient.subscriptions, "listLocations").returns(iterator);
    sandbox.stub(tools.tokenProvider.azureAccountProvider, "getSelectedSubscription").resolves({
      subscriptionId: "mockSubId",
      subscriptionName: "mockSubName",
      tenantId: "mockTID",
    });
    sandbox
      .stub(tools.tokenProvider.azureAccountProvider, "getIdentityCredentialAsync")
      .resolves(new MyTokenCredential());
    const res = await resourceGroupHelper.getLocations(
      tools.tokenProvider.azureAccountProvider,
      rmClient
    );
    assert.isTrue(res.isErr());
    if (res.isErr()) {
      assert.isTrue(res.error instanceof ListResourceGroupLocationsError);
    }
  });
});
