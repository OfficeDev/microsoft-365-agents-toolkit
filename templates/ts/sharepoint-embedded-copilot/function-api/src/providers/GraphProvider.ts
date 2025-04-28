import * as Graph from "@microsoft/microsoft-graph-client";
import axios, { AxiosRequestConfig, AxiosResponse } from "axios";
import { DriveItem } from "@microsoft/microsoft-graph-types";
import { AuthProvider } from "./AuthProvider";
import {
  IContainer,
  IContainerClientCreateRequest,
  IContainerColumn,
  IContainerServerCreateRequest,
  IContainerUpdateRequest,
} from "../../../common/schemas/ContainerSchemas";
import { Readable } from "stream";

export interface DriveItemColumnFilter {
  columnName: string;
  columnValue: string;
}

export interface IDriveProcessingItem {
  id: string;
  name: string;
  "@microsoft.graph.downloadUrl": string;
}

export interface IDriveItemFields {
  [key: string]: any;
}

export class GraphProvider {
  private _client: Graph.Client;
  private _authProvider: AuthProvider;
  private _containerTypeId: string = process.env.SPE_CONTAINER_TYPE_ID!;

  public constructor(authProvider: AuthProvider) {
    this._authProvider = authProvider;
    this._client = Graph.Client.init({
      authProvider: authProvider.getAuthHandler(),
    });
  }

  public async listContainers(): Promise<IContainer[]> {
    const response = await this._client
      .api("/storage/fileStorage/containers")
      .version("beta")
      .filter(`containerTypeId eq ${this._containerTypeId}`)
      .get();
    return response.value as IContainer[];
  }

  public async createContainer(
    clientCreateRequest: IContainerClientCreateRequest
  ): Promise<IContainer> {
    const createRequest: IContainerServerCreateRequest = {
      ...clientCreateRequest,
      containerTypeId: this._containerTypeId,
    };
    const response = await this._client
      .api("/storage/fileStorage/containers")
      .version("beta")
      .post(createRequest);
    return response as IContainer;
  }

  public async getContainer(id: string, loadColumns = true): Promise<IContainer> {
    const query = {
      $select:
        "id,displayName,containerTypeId,status,createdDateTime,description,customProperties,storageUsedInBytes,itemMajorVersionLimit,isItemVersioningEnabled",
      $expand: "permissions,drive",
    };
    const response = await this._client
      .api(`/storage/fileStorage/containers/${id}`)
      .query(query)
      .version("beta")
      .get();
    if (loadColumns) {
      response.columns = await this.getContainerColumns(id);
    }
    return response as IContainer;
  }

  public async deleteContainer(id: string): Promise<void> {
    const response = await this._client
      .api(`/storage/fileStorage/containers/${id}`)
      .version("beta")
      .delete();
  }
  public async updateContainer(
    id: string,
    updateRequest: IContainerUpdateRequest
  ): Promise<IContainer> {
    const response = await this._client
      .api(`/storage/fileStorage/containers/${id}`)
      .version("beta")
      .patch(updateRequest);
    return response as IContainer;
  }

  public async getContainerColumns(containerId: string): Promise<IContainerColumn[]> {
    const query = {
      //$select: "id,name,displayName,description,indexed,text,boolean,dateTime,currency,choice,hyperlinkOrPicture,number,personOrGroup",
      $filter: "readOnly eq false AND isDeletable eq true",
    };
    const response = await this._client
      .api(`/storage/fileStorage/containers/${containerId}/columns`)
      .version("beta")
      .query(query)
      .get();
    return response.value as IContainerColumn[];
  }

  public async addContainerColumn(
    containerId: string,
    column: IContainerColumn
  ): Promise<IContainerColumn> {
    const response = await this._client
      .api(`/storage/fileStorage/containers/${containerId}/columns`)
      .version("beta")
      .post(column);
    return response as IContainerColumn;
  }

  public async setContainerCustomProperties(
    containerId: string,
    customProperties: { [key: string]: any }
  ): Promise<void> {
    await this._client
      .api(`/storage/fileStorage/containers/${containerId}/customProperties`)
      .version("beta")
      .patch(customProperties);
  }

  public async getRootSiteUrl(): Promise<string> {
    const response = await this._client.api("/sites/root").get();
    return response.webUrl as string;
  }

  public async removeDriveSubscriptions(driveId: string): Promise<void> {
    interface ISubscription {
      id: string;
      resource: string;
    }

    const response = await this._client.api(`/subscriptions`).get();
    const subs = response.value as ISubscription[];
    //console.log(subs);
    for (const sub of subs) {
      if (sub.resource.includes(driveId)) {
        await this.removeSubscription(sub.id);
      }
    }
  }

  public async removeSubscription(subscriptionId: string): Promise<void> {
    console.log(await this._client.api(`/subscriptions/${subscriptionId}`).delete());
  }

  public async subscribeToDriveChanges(driveId: string, notificationUrl: string): Promise<any> {
    const now = new Date();
    const duration = 1000 * 60 * 4230; // max lifespan of driveItem subscription is 4230 minutes
    const expiry = new Date(now.getTime() + duration);
    const expiryDateTime = expiry.toISOString();

    const response = await this._client
      .api(`/subscriptions`)
      .version("beta")
      .post({
        changeType: "updated",
        notificationUrl: notificationUrl,
        resource: `/drives/${driveId}/root`,
        expirationDateTime: expiryDateTime,
        clientState: "",
      });
    return response as any;
  }

  public async getUnprocessedItems(
    driveId: string,
    filter: string
  ): Promise<IDriveProcessingItem[]> {
    const query = {
      $filter: `${filter}`,
    };
    const response = await this._client.api(`/drives/${driveId}/items`).query(query).get();
    return response.value as IDriveProcessingItem[];
  }

  public async setDriveItemFields(
    driveId: string,
    itemId: string,
    fields: IDriveItemFields
  ): Promise<void> {
    await this._client.api(`/drives/${driveId}/items/${itemId}/listitem/fields`).patch(fields);
  }

  public async getDriveItemByPath(driveId: string, itemPath: string): Promise<DriveItem> {
    const response = await this._client.api(`/drives/${driveId}/root:/${itemPath}`).get();
    return response as DriveItem;
  }

  public async createDriveItemAtRoot(driveId: string, itemName: string): Promise<DriveItem> {
    const response = await this._client
      .api(`/drives/${driveId}/root:/${itemName}:/content`)
      .putStream(null);
    return response as DriveItem;
  }

  public async getDriveItemStream(downloadUrl: string): Promise<Readable> {
    const token = await this._authProvider.getToken();
    const config: AxiosRequestConfig = {
      method: "get",
      url: downloadUrl,
      headers: {
        Authorization: `Bearer ${token}`,
      },
      responseType: "stream",
    };
    const response = await axios.get<Readable>(downloadUrl, config);
    return response.data;
  }
}
