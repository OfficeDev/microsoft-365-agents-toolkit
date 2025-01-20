// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { M365TokenProvider } from "@microsoft/teamsfx-api";
import axios, { AxiosInstance } from "axios";
import MockM365TokenProvider from "@microsoft/teamsapp-cli/src/commonlib/m365LoginUserPassword";
import AdmZip from "adm-zip";
import FormData from "form-data";
import fs from "fs-extra";

const sideloadingServiceEndpoint =
  process.env.SIDELOADING_SERVICE_ENDPOINT ??
  "{{SERVICE_ENDPOINT_PLACEHOLDER}}";
const sideloadingServiceScope =
  process.env.SIDELOADING_SERVICE_SCOPE ?? "{{SERVICE_SCOPE_PLACEHOLDER}}";

function delay(ms: number) {
  // tslint:disable-next-line no-string-based-set-timeout
  return new Promise((resolve) => setTimeout(resolve, ms));
}

export class M365TitleHelper {
  private static instance: M365TitleHelper;

  private axios: AxiosInstance;

  private constructor(serviceURL: string, access: string) {
    this.axios = axios.create({
      baseURL: serviceURL,
      headers: {
        authorization: `Bearer ${access}`,
        ConsistencyLevel: "eventual",
        "content-type": "application/json",
      },
    });
  }

  public static async init(
    endpoint: string = sideloadingServiceEndpoint,
    scope: string = sideloadingServiceScope,
    provider: M365TokenProvider = MockM365TokenProvider
  ): Promise<M365TitleHelper> {
    if (!M365TitleHelper.instance) {
      const res = await provider.getAccessToken({
        scopes: [scope],
      });
      if (res.isErr()) {
        throw res.error;
      }
      try {
        const envInfo = await axios.get("/config/v1/environment", {
          baseURL: endpoint,
          headers: {
            Authorization: `Bearer ${res.value}`,
          },
        });
        this.instance = new M365TitleHelper(
          envInfo.data.titlesServiceUrl,
          res.value
        );
      } catch (error: any) {
        throw error;
      }
    }
    return this.instance;
  }

  public async unacquire(id: string, retryTimes = 5) {
    if (!id) {
      return Promise.resolve(true);
    }
    return new Promise<boolean>(async (resolve) => {
      for (let i = 0; i < retryTimes; ++i) {
        try {
          await this.axios!.delete(`/catalog/v1/users/acquisitions/${id}`);
          console.info(`[Success] delete the M365 Title id: ${id}`);
          return resolve(true);
        } catch {
          await delay(2000);
        }
      }
      console.error(`[Failed] delete the M365 Title with id: ${id}`);
      return resolve(false);
    });
  }

  private checkZip(path: string) {
    const zip = new AdmZip(path, {});
    zip.getEntries();
  }

  public async acquire(packageFile: string): Promise<[string, string]> {
    try {
      this.checkZip(packageFile);
      const data = (await fs.readFile(packageFile)) as Buffer;
      const content = new FormData();
      content.append("package", data);
      const uploadResponse = await this.axios!.post(
        "/dev/v1/users/packages",
        content.getBuffer()
      );

      const operationId = uploadResponse.data.operationId;
      console.debug(`Package uploaded. OperationId: ${operationId as string}`);
      console.debug("Acquiring package ...");
      const acquireResponse = await this.axios!.post(
        "/dev/v1/users/packages/acquisitions",
        {
          operationId: operationId,
        }
      );
      const statusId = acquireResponse.data.statusId;
      console.debug(
        `Acquiring package with statusId: ${statusId as string} ...`
      );
      do {
        const statusResponse = await this.axios!.get(
          `/dev/v1/users/packages/status/${statusId as string}`
        );
        const resCode = statusResponse.status;
        console.debug(`Package status: ${resCode} ...`);
        if (resCode === 200) {
          const titleId: string = statusResponse.data.titleId;
          const appId: string = statusResponse.data.appId;
          console.info(`TitleId: ${titleId}`);
          console.info(`AppId: ${appId}`);
          console.info("Sideloading done.");
          return [titleId, appId];
        } else {
          await delay(2000);
        }
      } while (true);
    } catch (error) {
      if (error.response) {
        throw this.convertError(error);
      }
      throw error;
    }
  }
  private convertError(error: any): any {
    // add error details and trace to message
    const tracingId = (error.response.headers?.traceresponse ?? "") as string;
    const originalMessage = error.message as string;
    const innerError = error.response.data?.error || { code: "", message: "" };
    const finalMessage = `${originalMessage} (tracingId: ${tracingId}) ${
      innerError.code as string
    }: ${innerError.message as string} `;
    return {
      status: error.response?.status,
      tracingId,
      message: finalMessage,
    };
  }
}
