import { IGraph, Providers } from "@microsoft/mgt-element";
import { DriveItem } from "@microsoft/microsoft-graph-types";
import { DriveItemArrayConstructor, DriveItemConstructor, IDriveItem } from "../common/FileSchemas";
import * as Graph from "@microsoft/microsoft-graph-client";
import { GraphAuthProvider } from "./GraphAuthProvider";

export class GraphProvider {
  public static readonly instance: GraphProvider = new GraphProvider();
  private _authProvider: GraphAuthProvider = GraphAuthProvider.instance;
  private _client: Graph.Client;
  private get _providerClient(): IGraph | undefined {
    return Providers.globalProvider?.graph;
  }

  private constructor() {
    this._client = Graph.Client.init({
      authProvider: (done) => {
        this._authProvider
          .getToken()
          .then((token) => done(null, token))
          .catch((err) => done(err, null));
      },
    });
  }

  public async listItems(driveId: string, parentId = "root"): Promise<IDriveItem[]> {
    const endpoint = `/drives/${driveId}/items/${parentId}/children`;
    const query = {
      $expand: "listitem($expand=fields)",
      $select:
        "id,name,createdDateTime,lastModifiedBy,lastModifiedDateTime,size,folder,file,root,parentReference,webUrl,webDavUrl,content.downloadUrl",
    };
    const response = await this._providerClient?.api(endpoint).query(query).get();
    const items: DriveItem[] = response.value as DriveItem[];
    return DriveItemArrayConstructor.from(items);
  }

  public async getItem(driveId: string, itemId: string): Promise<IDriveItem> {
    const endpoint = `/drives/${driveId}/items/${itemId}`;
    const query = {
      $expand: "listitem($expand=fields)",
      $select:
        "id,name,createdDateTime,lastModifiedBy,lastModifiedDateTime,size,folder,file,root,parentReference,webUrl,webDavUrl,content.downloadUrl",
    };
    const response = await this._providerClient?.api(endpoint).query(query).get();
    const item: DriveItem = response as DriveItem;
    return DriveItemConstructor.from(item);
  }

  public async getItemPath(item: IDriveItem): Promise<IDriveItem[]> {
    if (!item || item.root) {
      return [];
    }
    if (!item.parentReference) {
      return [item];
    }
    let parentPath: IDriveItem[] = [];
    try {
      const parent = await this.getItem(item.parentReference.driveId!, item.parentReference.id!);
      parentPath = await this.getItemPath(parent);
    } catch (e) {}
    return [...parentPath, item];
  }

  public async uploadFile(driveId: string, file: File, parentId = "root"): Promise<IDriveItem> {
    const smallSizeMax = 4 * 1024 * 1024;
    if (file.size > smallSizeMax) {
      return this._uploadLargeFile(driveId, file, parentId);
    } else {
      return this._uploadSmallFile(driveId, file, parentId);
    }
  }

  private async _uploadLargeFile(
    driveId: string,
    file: File,
    parentId: string
  ): Promise<IDriveItem> {
    const options: Graph.LargeFileUploadTaskOptions = {
      // Chunk size must be a multiple of 320 KiB: https://learn.microsoft.com/en-us/onedrive/developer/rest-api/api/driveitem_createuploadsession?view=odsp-graph-online#upload-bytes-to-the-upload-session
      rangeSize: 10 * 320 * 1024, // 3.2 MiB
      uploadEventHandlers: {
        progress: (range, _) => {
          console.log(`Uploaded bytes ${range?.minValue}-${range?.maxValue} of ${file.size}`);
        },
      },
    };
    const endpoint = `/drives/${driveId}/items/${parentId}:/${file.name}:/createUploadSession`;
    const payload = {
      item: {
        "@microsoft.graph.conflictBehavior": "rename",
        name: file.name,
      },
    };
    const session = await Graph.LargeFileUploadTask.createUploadSession(
      this._client,
      endpoint,
      payload
    );
    const upload = new Graph.FileUpload(file, file.name, file.size);
    const task = new Graph.LargeFileUploadTask(this._client, upload, session, options);
    const result = await task.upload();
    console.log(result);
    return result.responseBody as IDriveItem;
  }

  private async _uploadSmallFile(
    driveId: string,
    file: File,
    parentId: string
  ): Promise<IDriveItem> {
    const fileReader = new FileReader();
    fileReader.readAsArrayBuffer(file);
    return new Promise<IDriveItem>((resolve, reject) => {
      fileReader.addEventListener("loadend", async (event: any) => {
        const endpoint = `/drives/${driveId}/items/${parentId}:/${file.name}:/content`;
        this._providerClient
          ?.api(endpoint)
          .putStream(fileReader.result)
          .then(async (response) => {
            resolve(response as IDriveItem);
          })
          .catch((error) => {
            reject(new Error(`Failed to upload file ${file.name}: ${error.message}`));
          });
      });
      fileReader.addEventListener("error", (event: any) => {
        reject(new Error(`Error on reading file: ${event.message}`));
      });
    });
  }

  public async newDocument(
    driveId: string,
    parentId: string,
    extension: string
  ): Promise<IDriveItem> {
    const endpoint = `/drives/${driveId}/items/${parentId}/children`;
    const body = {
      name: `New Document.${extension}`,
      file: {},
      "@microsoft.graph.conflictBehavior": "rename",
    };
    return (await this._providerClient?.api(endpoint).post(body)) as IDriveItem;
  }

  public async createFolder(driveId: string, parentId: string, newFolderName: string) {
    const endpoint = `/drives/${driveId}/items/${parentId}/children`;
    const body = {
      name: newFolderName,
      folder: {},
      "@microsoft.graph.conflictBehavior": "rename",
    };
    return this._providerClient?.api(endpoint).post(body);
  }

  public async renameItem(driveId: string, itemId: string, newName: string) {
    const endpoint = `/drives/${driveId}/items/${itemId}`;
    const body = {
      name: newName,
    };
    return this._providerClient?.api(endpoint).patch(body);
  }

  public async deleteItem(driveId: string, itemId: string) {
    const endpoint = `/drives/${driveId}/items/${itemId}`;
    return this._providerClient?.api(endpoint).delete();
  }

  public async getPath(driveId: string, itemId: string): Promise<string> {
    const endpoint = `/drives/${driveId}/items/${itemId}`;
    const response = await this._providerClient?.api(endpoint).get();
    if (response.root !== undefined) {
      return endpoint;
    }
    return `${response.parentReference.path}/${response.name}`;
  }

  public async getPreviewUrl(driveId: string, itemId: string): Promise<URL> {
    const endpoint = `/drives/${driveId}/items/${itemId}/preview`;
    const response = await this._providerClient?.api(endpoint).post({});
    const url = new URL(response.getUrl);
    url.searchParams.set("nb", "true");
    return url;
  }

  public async getSocketUrl(driveId: string): Promise<URL> {
    const endpoint = `/drives/${driveId}/root/subscriptions/socketIo`;
    const response = await this._providerClient?.api(endpoint).get();
    const url = new URL(response.notificationUrl as string);
    return url;
  }

  public async getSpUrl(): Promise<string> {
    const endpoint = `/sites/root`;
    const response = await this._providerClient?.api(endpoint).get();
    return response.webUrl;
  }

  public async getPdfUrl(driveId: string, itemId: string): Promise<URL> {
    const endpoint = `/drives/${driveId}/items/${itemId}/content?format=pdf`;
    // NB: We use the `raw` response type so that we can capture the redirected URL of the PDF file.
    // See: https://learn.microsoft.com/en-us/graph/api/driveitem-get-content-format?view=graph-rest-1.0&tabs=http#response
    const response = await this._providerClient
      ?.api(endpoint)
      .responseType(Graph.ResponseType.RAW)
      .get();
    const url = new URL(response.url);
    return url;
  }

  public async getContentStream(
    driveId: string,
    itemId: string
  ): Promise<ReadableStream<Uint8Array>> {
    const endpoint = `/drives/${driveId}/items/${itemId}/content`;
    return await this._providerClient?.api(endpoint).getStream();
  }
}
