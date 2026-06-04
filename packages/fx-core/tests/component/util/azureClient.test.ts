import chai from "chai";
import chaiAsPromised from "chai-as-promised";
import * as sinon from "sinon";
import { azureClientHelper } from "../../../src/component/utils/azureClient";
import { InvalidAzureCredentialError } from "../../../src/error";
import { MockedAzureAccountProvider } from "../../core/utils";

chai.use(chaiAsPromised);

describe("azureClient test", () => {
  const sandbox = sinon.createSandbox();

  beforeEach(async () => {});

  afterEach(async () => {
    sandbox.restore();
  });

  it("getChallengeHandler returns invalid token", async () => {
    const tokenProvider = new MockedAzureAccountProvider();
    sandbox.stub(tokenProvider, "getIdentityCredentialAsync").resolves(undefined);

    const getTokenForChallenge = azureClientHelper.getChallengeHandler(tokenProvider);
    try {
      await getTokenForChallenge({
        wwwAuthenticate: "faked-claim",
        scopes: ["https://management.azure.com/.default"],
      });
      chai.assert.fail("Expected getTokenForChallenge to reject");
    } catch (e) {
      chai.assert.isTrue(e instanceof InvalidAzureCredentialError);
    }
  });

  it("getChallengeHandler happy pass", async () => {
    const tokenProvider = new MockedAzureAccountProvider();
    sandbox.stub(tokenProvider, "getIdentityCredentialAsync").resolves({
      getToken: async function (scopes: string | string[]) {
        return { token: "fake-token", expiresOnTimestamp: 0 };
      },
    });

    const getTokenForChallenge = azureClientHelper.getChallengeHandler(tokenProvider);
    const token = await getTokenForChallenge({
      wwwAuthenticate: "faked-claim",
      scopes: ["https://management.azure.com/.default"],
    });
    chai.assert.equal(token, "fake-token");
  });
});
