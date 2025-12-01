// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import axios from "axios";
import * as chai from "chai";
import { EnvConstants } from "./constants";

interface IStaticSiteObject {
  endpoint: string;
}

export class StaticSiteValidator {
  public static init(context: any): IStaticSiteObject {
    const endpoint = this.getStaticSiteEndpointFromContext(context);
    return { endpoint };
  }

  public static async validateDeploy(
    staticSite: IStaticSiteObject
  ): Promise<void> {
    const tabEndpoint = `${staticSite.endpoint}/index.html`;
    console.log(`getting tab endpoint ${tabEndpoint}`);
    const response = await axios.get(tabEndpoint);
    chai.assert.equal(response.status, 200);
  }

  private static getStaticSiteEndpointFromContext(context: any): string {
    console.log("context: ", JSON.stringify(context));
    return context[EnvConstants.APP_ENDPOINT];
  }
}
