// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { err, FxError, ok, Result, SystemError } from "@microsoft/teamsfx-api";
import { WrappedAxiosClient } from "../common/wrappedAxiosClient";
import { hooks } from "@feathersjs/hooks";
import { ErrorContextMW } from "../common/globalVars";
import { getDefaultString } from "../common/localizeUtils";

export const listSensitivityLabelScope = "InformationProtectionPolicy.Read";

const graphAPIEndpoint = "https://graph.microsoft.com";
const listSensitivityLabelAPIPath = "/beta/me/informationProtection/sensitivityLabels";
const errorSourceName = "GraphAPI";
const GeneralLabelDisplayName = "General";

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
    mock = false
  ): Promise<Result<SensitivityLabel[], FxError>> {
    if (mock) {
      return ok([
        {
          id: "87ba5c36-b7cf-4793-bbc2-bd5b3a9f95ca",
          name: "Personal",
          description: undefined,
          displayName: "Personal",
        },
        {
          id: "87867195-f2b8-4ac2-b0b6-6bb73cb33afc",
          name: "Not Restricted",
          description: undefined,
          displayName: "Public",
        },
        {
          id: "f42aa342-8706-4288-bd11-ebb85995028c",
          name: "Internal",
          description: "Billjo - Removed footer from label",
          displayName: "General",
        },
      ]);
    }
    try {
      const requester = WrappedAxiosClient.create({
        baseURL: graphAPIEndpoint,
      });
      requester.defaults.headers.common["Authorization"] = `Bearer ${token}`;
      requester.defaults.headers.common["Content-Type"] = "application/json";

      const response = await RetryHandler.Retry(() => requester.get(listSensitivityLabelAPIPath));

      if (response && response.data && response.data.value) {
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

  async getGeneralSentivityLabelId(token: string, mock = false): Promise<Result<string, FxError>> {
    const result = await this.listSensitivityLabels(token, mock);
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
