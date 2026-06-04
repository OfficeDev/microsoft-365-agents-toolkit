import chai from "chai";
import chaiAsPromised from "chai-as-promised";
import * as sinon from "sinon";
import { MockedAzureAccountProvider } from "../../core/utils";
import { azureClientHelper } from "../../../src/component/utils/azureClient";
import { InvalidAzureCredentialError } from "../../../src/error";
import { BearerChallengePolicy } from "../../../src/component/utils/pipelinePolicy";
import { get, head } from "lodash";

chai.use(chaiAsPromised);

describe("BearerChallengePolicy test", () => {
  const sandbox = sinon.createSandbox();

  beforeEach(async () => {});

  afterEach(async () => {
    sandbox.restore();
  });

  it("returns 401 error and trigger challenge retrieval", async () => {
    const challengePolicy = new BearerChallengePolicy(async (scopes) => {
      return "fake-token";
    });

    const next = async (request: any) => {
      return {
        status: 401,
        headers: {
          get: (key: string) => {
            return key === "WWW-Authenticate" ? "faked-claim" : undefined;
          },
        },
      };
    };
    const request = {
      headers: {
        map: new Map<string, string>(),
        get(key: string) {
          return this.map.get(key);
        },
        set(key: string, value: string) {
          this.map.set(key, value);
        },
      },
    };
    const response = await challengePolicy.sendRequest(request as any, next as any);
    chai.assert.equal(response.status, 401);
  });

  it("returns 200", async () => {
    const challengePolicy = new BearerChallengePolicy(async (scopes) => {
      return "fake-token";
    });

    const next = async (request: any) => {
      return {
        status: 200,
        headers: {
          get: (key: string) => {
            return key === "Authorization" ? "Bearer fake-token" : undefined;
          },
        },
      };
    };
    const request = {
      headers: {
        map: new Map<string, string>(),
        get(key: string) {
          return this.map.get(key);
        },
        set(key: string, value: string) {
          this.map.set(key, value);
        },
      },
    };
    const response = await challengePolicy.sendRequest(request as any, next as any);
    chai.assert.equal(response.status, 200);
  });
});
