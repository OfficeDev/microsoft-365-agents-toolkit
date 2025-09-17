// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import {
  PipelinePolicy,
  PipelineRequest,
  SendRequest,
  PipelineResponse,
} from "@azure/core-rest-pipeline";
import { AuthenticationWWWAuthenticateRequest } from "@microsoft/teamsfx-api";
import { AzureScopes } from "../../common/constants";

export class BearerChallengePolicy implements PipelinePolicy {
  public readonly name = "BearerChallengePolicy";
  private readonly challengeRetryHeader = "x-atk-challenge-retry";

  public constructor(
    private readonly getTokenForChallenge: (
      scopes: AuthenticationWWWAuthenticateRequest
    ) => Promise<string | undefined>
  ) {}

  public async sendRequest(request: PipelineRequest, next: SendRequest): Promise<PipelineResponse> {
    const initial = await next(request);

    // Only attempt a single retry on auth challenges
    if (initial.status === 401 && !request.headers.get(this.challengeRetryHeader)) {
      const header =
        initial.headers.get("WWW-Authenticate") || initial.headers.get("www-authenticate");
      if (header) {
        request.headers.set(this.challengeRetryHeader, "1");

        const token = await this.getTokenForChallenge({
          wwwAuthenticate: header,
          scopes: AzureScopes,
        });
        if (token) {
          request.headers.set("Authorization", `Bearer ${token}`);
          return await next(request);
        }
      }
    }

    return initial;
  }
}
