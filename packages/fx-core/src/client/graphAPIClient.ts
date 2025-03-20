// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { err, FxError, ok, Result, SystemError } from "@microsoft/teamsfx-api";
import { WrappedAxiosClient } from "../common/wrappedAxiosClient";
import { hooks } from "@feathersjs/hooks";
import { ErrorContextMW } from "../common/globalVars";
import { getDefaultString } from "../common/localizeUtils";
import { globalStateGet, globalStateUpdate } from "../common/globalState";

export const listSensitivityLabelScope = "InformationProtectionPolicy.Read";

const graphAPIEndpoint = "https://graph.microsoft.com";
const listSensitivityLabelAPIPath = "/beta/me/informationProtection/sensitivityLabels";
const errorSourceName = "GraphAPI";
const GeneralLabelDisplayName = "General";
const listSensitivityLabelCacheKey = "listSensitivityLabelCacheKey";

export class SensitivityLabel {
  id?: string;
  name?: string;
  description?: string;
  displayName?: string;
}

export class RetryHandler {
  public static RETRIES = 3;
  public static async Retry<T>(fn: () => Promise<T>): Promise<T | undefined> {
    let retries = this.RETRIES;
    let lastError: any;
    while (retries > 0) {
      retries--;
      try {
        return await fn();
      } catch (e: any) {
        lastError = e;
        if (retries > 0) {
          await new Promise((resolve) => setTimeout(resolve, 5000));
        }
      }
    }
    throw lastError;
  }
}

export class GraphAPIClient {
  @hooks([ErrorContextMW({ source: "Graph", component: "GraphAPIClient" })])
  async listSensitivityLabels(
    token: string,
    useCache = false,
    accountUniqueName = ""
  ): Promise<Result<SensitivityLabel[], FxError>> {
    try {
      if (useCache) {
        const cacheKey = `${listSensitivityLabelCacheKey}-${accountUniqueName}`;
        const cacheValueRes = await globalStateGet(cacheKey);
        if (cacheValueRes.isOk()) {
          const timeStamp = cacheValueRes.value.unixTimestamp;
          // if cache data is within 1 days, use the cache.
          if (Date.now() - timeStamp < 1000 * 60 * 60 * 24) {
            return ok(cacheValueRes.value.labels);
          }
        }
      }
      const requester = WrappedAxiosClient.create({
        baseURL: graphAPIEndpoint,
      });
      requester.defaults.headers.common["Authorization"] = `Bearer ${token}`;
      requester.defaults.headers.common["Content-Type"] = "application/json";

      const response = await RetryHandler.Retry(() => requester.get(listSensitivityLabelAPIPath));

      if (response && response.data && response.data.value) {
        if (useCache) {
          const cacheKey = `${listSensitivityLabelCacheKey}-${accountUniqueName}`;
          const cacheValue: ListSensitivityCacheValue = {
            labels: response.data.value,
            unixTimestamp: Date.now(),
          };
          await globalStateUpdate(cacheKey, cacheValue);
        }
        return ok(response.data.value);
      } else {
        return err(
          new SystemError({
            name: "listSensitivityLabelsError",
            message: getDefaultString(
              "error.graphAPI.apiFailed.message",
              "listSensitivityLabels",
              "empty data"
            ),
            source: errorSourceName,
          })
        );
      }
    } catch (error: any) {
      return err(
        new SystemError({
          name: "listSensitivityLabelsError",
          message: getDefaultString(
            "error.graphAPI.apiFailed.message",
            "listSensitivityLabels",
            error.message
          ),
          source: errorSourceName,
        })
      );
    }
  }

  async getGeneralSentivityLabelId(token: string): Promise<Result<string, FxError>> {
    const result = await this.listSensitivityLabels(token);
    if (result.isErr()) {
      return err(result.error);
    }
    const labels = result.value;
    const generalLabel = labels.find((label) => label.displayName === GeneralLabelDisplayName);
    if (generalLabel && generalLabel.id) {
      return ok(generalLabel.id);
    } else {
      return err(
        new SystemError({
          name: "getGeneralSentivityLabelIdError",
          message: getDefaultString(
            "error.graphAPI.apiFailed.message",
            "getGeneralSentivityLabelId",
            "General label not found"
          ),
          source: errorSourceName,
        })
      );
    }
  }
}

export const graphAPIClient = new GraphAPIClient();

interface ListSensitivityCacheValue {
  labels: SensitivityLabel[];
  unixTimestamp: number;
}
