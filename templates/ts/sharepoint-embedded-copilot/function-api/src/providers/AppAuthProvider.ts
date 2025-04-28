import { AuthProvider } from "./AuthProvider";

export class AppAuthProvider extends AuthProvider {
  public constructor(tid: string, host = "https://graph.microsoft.com") {
    if (host.charAt(host.length - 1) !== "/") {
      host += "/";
    }
    host += ".default";
    super(tid, [host]);
  }

  public async getToken(): Promise<string> {
    const result = await this.client.acquireTokenByClientCredential({
      scopes: this.scopes,
    });
    return result!.accessToken;
  }
}
