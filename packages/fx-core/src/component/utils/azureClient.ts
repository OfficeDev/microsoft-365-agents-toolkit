// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { AuthenticationWWWAuthenticateRequest, AzureAccountProvider } from "@microsoft/teamsfx-api";
import { ResourceManagementClient } from "@azure/arm-resources";
import { InvalidAzureCredentialError } from "../../error";
import { Pipeline, PipelinePolicy } from "@azure/core-rest-pipeline";
import { BearerChallengePolicy } from "./pipelinePolicy";
import { AzureScopes } from "../../common/constants";

class AzureClientHelper {
  async createRmClient(azureAccountProvider: AzureAccountProvider, subscriptionId: string) {
    const azureToken = await azureAccountProvider.getIdentityCredentialAsync();
    if (azureToken === undefined) {
      throw new InvalidAzureCredentialError();
    }
    await azureAccountProvider.setSubscription(subscriptionId);
    const rmClient = new ResourceManagementClient(azureToken, subscriptionId, {
      userAgentOptions: { userAgentPrefix: "AgentsToolkit" },
    });
    this.addPipelinePolicy(
      rmClient.pipeline,
      new BearerChallengePolicy(this.getChallengeHandler(azureAccountProvider))
    );
    return rmClient;
  }

  addPipelinePolicy(pipeline: Pipeline, policy: PipelinePolicy) {
    pipeline.addPolicy(policy, { phase: "Sign" });
  }

  getChallengeHandler = (tokenProvider: AzureAccountProvider) => {
    const getTokenForChallenge = async (
      scopes: AuthenticationWWWAuthenticateRequest
    ): Promise<string> => {
      const azureToken = await tokenProvider.getIdentityCredentialAsync(false, scopes);
      if (!azureToken) {
        throw new InvalidAzureCredentialError();
      }
      const token = (await azureToken.getToken(AzureScopes)) as { token: string };
      return token.token;
    };

    return getTokenForChallenge;
  };
}

export const azureClientHelper = new AzureClientHelper();
