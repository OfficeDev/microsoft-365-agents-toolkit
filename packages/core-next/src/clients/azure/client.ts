// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { AxiosInstance } from "axios";
import { Result, ok, err } from "neverthrow";
import { AtkContext } from "../../core/context";
import { AtkError, systemError, userError } from "../../core/error";
import { createHttpClient } from "../../http/httpClient";
import { sendWithRetry } from "../../http/retry";
import {
  ARM_BASE_URL,
  ARM_DEPLOYMENT_API_VERSION,
  ARM_WEBAPPS_API_VERSION,
  ArmDeploymentRequest,
  ArmDeploymentResponse,
  isTerminalState,
} from "./types";

const SOURCE = "AzureArmClient";
const DEPLOYMENT_POLL_INTERVAL_MS = 10_000; // 10 seconds
const DEPLOYMENT_POLL_MAX_ATTEMPTS = 120; // ~20 minutes

/**
 * Azure Resource Manager HTTP client.
 * Uses the ARM REST API directly via axios — no heavy Azure SDK dependency.
 */
export class AzureArmClient {
  private readonly http: AxiosInstance;

  constructor(ctx: AtkContext, token: string) {
    this.http = createHttpClient(ctx, { baseURL: ARM_BASE_URL });
    this.http.defaults.headers.common["Authorization"] = `Bearer ${token}`;
    this.http.defaults.headers.common["Content-Type"] = "application/json";
  }

  /**
   * Deploy an ARM template and wait for completion.
   * PUT /subscriptions/{sub}/resourceGroups/{rg}/providers/Microsoft.Resources/deployments/{name}
   */
  async deployTemplate(
    subscriptionId: string,
    resourceGroupName: string,
    deploymentName: string,
    body: ArmDeploymentRequest
  ): Promise<Result<ArmDeploymentResponse, AtkError>> {
    const path = `/subscriptions/${subscriptionId}/resourceGroups/${resourceGroupName}/providers/Microsoft.Resources/deployments/${deploymentName}`;
    try {
      const res = await sendWithRetry(
        () =>
          this.http.put<ArmDeploymentResponse>(path, body, {
            params: { "api-version": ARM_DEPLOYMENT_API_VERSION },
          }),
        0 // no retries on PUT — ARM handles idempotency
      );

      const deployment = res.data;
      const state = deployment.properties?.provisioningState ?? "";

      // If already terminal, return immediately
      if (isTerminalState(state)) {
        return this.handleTerminalState(deployment, deploymentName, resourceGroupName);
      }

      // Otherwise poll until terminal
      return await this.pollDeployment(subscriptionId, resourceGroupName, deploymentName);
    } catch (e) {
      return err(this.wrapError("deployTemplate", e));
    }
  }

  /**
   * Poll a deployment until it reaches a terminal state.
   */
  private async pollDeployment(
    subscriptionId: string,
    resourceGroupName: string,
    deploymentName: string
  ): Promise<Result<ArmDeploymentResponse, AtkError>> {
    const path = `/subscriptions/${subscriptionId}/resourceGroups/${resourceGroupName}/providers/Microsoft.Resources/deployments/${deploymentName}`;

    for (let attempt = 0; attempt < DEPLOYMENT_POLL_MAX_ATTEMPTS; attempt++) {
      await delay(DEPLOYMENT_POLL_INTERVAL_MS);

      try {
        const res = await this.http.get<ArmDeploymentResponse>(path, {
          params: { "api-version": ARM_DEPLOYMENT_API_VERSION },
        });

        const state = res.data.properties?.provisioningState ?? "";
        if (isTerminalState(state)) {
          return this.handleTerminalState(res.data, deploymentName, resourceGroupName);
        }
      } catch (e) {
        return err(this.wrapError("pollDeployment", e));
      }
    }

    return err(
      systemError(
        "DeploymentPollTimeout",
        `ARM deployment '${deploymentName}' did not complete within polling timeout`,
        {
          source: SOURCE,
        }
      )
    );
  }

  /**
   * Check terminal provisioning state and return ok/err accordingly.
   */
  private handleTerminalState(
    deployment: ArmDeploymentResponse,
    deploymentName: string,
    resourceGroupName: string
  ): Result<ArmDeploymentResponse, AtkError> {
    const state = deployment.properties?.provisioningState ?? "";
    if (state === "Succeeded") {
      return ok(deployment);
    }
    const armErr = deployment.properties?.error;
    let message: string;
    if (armErr) {
      message = `ARM deployment '${deploymentName}' in resource group '${resourceGroupName}' failed: [${armErr.code}] ${armErr.message}`;
      // Append nested details so CI logs show which specific resource(s) failed.
      if (armErr.details?.length) {
        const detailLines = armErr.details.map((d) => `  - [${d.code}] ${d.message}`).join("\n");
        message += `\nDetails:\n${detailLines}`;
      }
    } else {
      message = `ARM deployment '${deploymentName}' in resource group '${resourceGroupName}' finished with state: ${state}`;
    }
    return err(userError("DeployArmError", message, { source: SOURCE }));
  }

  /**
   * Get the SCM hostname for an Azure Web App / Function App.
   * GET /subscriptions/{sub}/resourceGroups/{rg}/providers/Microsoft.Web/sites/{name}
   */
  async getScmEndpoint(
    subscriptionId: string,
    resourceGroupName: string,
    siteName: string
  ): Promise<Result<string, AtkError>> {
    const path = `/subscriptions/${subscriptionId}/resourceGroups/${resourceGroupName}/providers/Microsoft.Web/sites/${siteName}`;
    try {
      const res = await sendWithRetry(() =>
        this.http.get<{ properties?: { enabledHostNames?: string[] } }>(path, {
          params: { "api-version": ARM_WEBAPPS_API_VERSION },
        })
      );
      const hostNames = res.data.properties?.enabledHostNames ?? [];
      const scmHost = hostNames.find((h) => h.includes("scm"));
      if (!scmHost) {
        return err(
          userError(
            "ScmEndpointNotFound",
            `Cannot find SCM hostname for site '${siteName}'. Available: ${hostNames.join(", ")}`,
            {
              source: SOURCE,
            }
          )
        );
      }
      return ok(`https://${scmHost}`);
    } catch (e) {
      return err(this.wrapError("getScmEndpoint", e));
    }
  }

  /**
   * Restart a web app / function app.
   * POST /subscriptions/{sub}/resourceGroups/{rg}/providers/Microsoft.Web/sites/{name}/restart
   */
  async restartSite(
    subscriptionId: string,
    resourceGroupName: string,
    siteName: string
  ): Promise<Result<void, AtkError>> {
    const path = `/subscriptions/${subscriptionId}/resourceGroups/${resourceGroupName}/providers/Microsoft.Web/sites/${siteName}/restart`;
    try {
      await sendWithRetry(() =>
        this.http.post(path, null, {
          params: { "api-version": ARM_WEBAPPS_API_VERSION },
        })
      );
      return ok(undefined);
    } catch (e) {
      return err(this.wrapError("restartSite", e));
    }
  }

  private wrapError(apiName: string, e: unknown): AtkError {
    if (e && typeof e === "object" && "response" in e) {
      const axErr = e as {
        response?: { status?: number; data?: { error?: { code?: string; message?: string } } };
        message?: string;
      };
      const status = axErr.response?.status ?? 0;
      const code = axErr.response?.data?.error?.code ?? "ArmApiError";
      const msg = axErr.response?.data?.error?.message ?? axErr.message ?? String(e);
      if (status >= 400 && status < 500) {
        return userError(code, `[${apiName}] ${msg}`, { source: SOURCE });
      }
      return systemError(code, `[${apiName}] ${msg}`, { source: SOURCE });
    }
    // Surface inner errors from AggregateError (e.g. Node.js DNS/socket failures)
    const msg =
      e instanceof AggregateError
        ? `${e.message}: ${e.errors.map((inner) => String(inner)).join("; ")}`
        : String(e);
    return systemError("ArmApiError", `[${apiName}] ${msg}`, { source: SOURCE });
  }
}

function delay(ms: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, ms));
}
