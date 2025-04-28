import {
  IContainer,
  IContainerClientCreateRequest,
  IContainerUpdateRequest,
} from "../../../common/schemas/ContainerSchemas";
import * as Scopes from "../common/Scopes";
import { CustomAppApiAuthProvider } from "./CustomAppApiAuthProvider";

export class ContainersApiProvider {
  public readonly apiUrl: string =
    process.env.REACT_APP_SAMPLE_API_URL || "http://localhost:7071/api";
  public readonly apiScope: string = Scopes.SAMPLE_API_CONTAINER_MANAGE;

  public static readonly instance: ContainersApiProvider = new ContainersApiProvider();
  private _authProvider: { getToken: () => Promise<string> };
  public get authProvider() {
    return this._authProvider;
  }
  public set authProvider(value: { getToken: () => Promise<string> }) {
    this._authProvider = value;
  }

  private constructor() {
    this._authProvider = CustomAppApiAuthProvider.instance;
  }

  public async list(): Promise<IContainer[]> {
    const request: RequestInit = {
      method: "GET",
      headers: this._headers(await this.authProvider.getToken()),
    };
    return (await this._send("/containers", request)) as IContainer[];
  }

  public async get(id: string): Promise<IContainer> {
    const request: RequestInit = {
      method: "GET",
      headers: this._headers(await this.authProvider.getToken()),
    };
    return (await this._send(`/containers/${id}`, request)) as IContainer;
  }

  public async create(container: IContainerClientCreateRequest): Promise<IContainer> {
    const request: RequestInit = {
      method: "POST",
      headers: this._headers(await this.authProvider.getToken()),
      body: JSON.stringify(container),
    };
    return (await this._send("/containers", request)) as IContainer;
  }

  public async update(container: IContainer): Promise<IContainer> {
    const id = container.id;
    if (!id) {
      throw new Error("Container id is required");
    }
    const containerUpdate: IContainerUpdateRequest = container as IContainerUpdateRequest;
    delete containerUpdate.id;
    const request: RequestInit = {
      method: "PUT",
      headers: this._headers(await this.authProvider.getToken()),
      body: JSON.stringify(containerUpdate),
    };
    return (await this._send(`/containers/${id}`, request)) as IContainer;
  }

  public async registerContainerType(): Promise<any> {
    const request: RequestInit = {
      method: "PUT",
      headers: this._headers(await this.authProvider.getToken()),
    };
    return await this._send("/registerContainerType", request);
  }

  /** Private methods for requests **/

  private async _send(endpoint: string, request: RequestInit): Promise<any> {
    const response = await fetch(this._url(endpoint), request);
    if (!response.ok) {
      throw new Error(`Request failed: ${response.statusText}`);
    }
    return await response.json();
  }

  private _headers(token: string): HeadersInit {
    return {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
    };
  }

  private _url(endpoint: string): string {
    return `${this.apiUrl}${endpoint}`;
  }
}
