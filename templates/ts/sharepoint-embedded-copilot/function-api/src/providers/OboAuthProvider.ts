import { OnBehalfOfRequest } from "@azure/msal-node";
import { AuthProvider } from "./AuthProvider";
import { JwtProvider } from "./JwtProvider";

export class OboAuthProvider extends AuthProvider {
  public constructor(
    private _jwt: JwtProvider,
    scopes: string[] = ["FileStorageContainer.Selected"]
  ) {
    super(_jwt.tid!, scopes);
  }
  public async getToken(): Promise<string> {
    const request: OnBehalfOfRequest = {
      oboAssertion: this._jwt.token,
      scopes: this.scopes,
    };
    const result = await this.client.acquireTokenOnBehalfOf(request);
    return result!.accessToken;
  }
}
