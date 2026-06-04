import { ListAccountSasResponse, StorageAccounts } from "@azure/arm-storage";
import chai from "chai";
import chaiAsPromised from "chai-as-promised";
import * as sinon from "sinon";
import * as tools from "../../../src/common/utils";
import {
  generateSasToken,
  getAzureAccountCredential,
} from "../../../src/component/utils/azureResourceOperation";
import { MockedAzureAccountProvider } from "../../core/utils";

chai.use(chaiAsPromised);

describe("Azure Resource Operation test", () => {
  const sandbox = sinon.createSandbox();

  beforeEach(async () => {
    sandbox.stub(tools, "waitSeconds").resolves();
  });

  afterEach(async () => {
    sandbox.restore();
  });

  it("should get Azure account credential error", async () => {
    const tokenProvider = new MockedAzureAccountProvider();
    sandbox.stub(tokenProvider, "getIdentityCredentialAsync").resolves(undefined);
    await chai.expect(getAzureAccountCredential(tokenProvider)).to.be.eventually.rejectedWith("");
  });

  it("should generate Sas token error", async () => {
    const storageAccounts = {
      listAccountSAS: async function (): Promise<ListAccountSasResponse> {
        return {
          accountSasToken: "abc",
        };
      },
    } as unknown as StorageAccounts;
    sandbox.stub(storageAccounts, "listAccountSAS").throws(new Error("error"));
    await chai
      .expect(generateSasToken(storageAccounts, "test", "test"))
      .to.be.eventually.rejectedWith("");
  });

  it("should generate Sas token with empty response", async () => {
    const storageAccounts = {
      listAccountSAS: async function (): Promise<ListAccountSasResponse> {
        return {
          accountSasToken: "",
        };
      },
    } as unknown as StorageAccounts;
    await chai
      .expect(generateSasToken(storageAccounts, "test", "test"))
      .to.be.eventually.rejectedWith("");
  });
});
