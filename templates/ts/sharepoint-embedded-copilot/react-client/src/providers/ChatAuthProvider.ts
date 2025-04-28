import { InteractionRequiredAuthError, PublicClientApplication } from "@azure/msal-browser";
import * as Constants from "../common/Constants";
import { GraphProvider } from "./GraphProvider";

export class ChatAuthProvider {
  private static _instance?: ChatAuthProvider;
  private _client: PublicClientApplication;

  public static async getInstance(): Promise<ChatAuthProvider> {
    const spHostname = await GraphProvider.instance.getSpUrl();
    if (!ChatAuthProvider._instance) {
      const spHostname = await GraphProvider.instance.getSpUrl();
      ChatAuthProvider._instance = new ChatAuthProvider(spHostname);
      await ChatAuthProvider._instance.initialize();
    }
    return ChatAuthProvider._instance;
  }

  private constructor(public readonly hostname: string) {
    this._client = new PublicClientApplication({
      auth: {
        clientId: Constants.AZURE_CLIENT_ID!,
        authority: Constants.AUTH_AUTHORITY,
        redirectUri: window.location.origin,
      },
      cache: {
        // https://github.com/AzureAD/microsoft-authentication-library-for-js/blob/dev/lib/msal-browser/docs/caching.md
        /*
                Cache Location	| Cleared on |	Shared between windows/tabs |	Redirect flow supported
                -----------------   ----------  -------------------------   ------------------------
                sessionStorage |	window/tab close |	No |	Yes
                localStorage |	browser close | Yes |   Yes
                memoryStorage | page |  refresh/navigation | No |	No
                */
        cacheLocation: "localStorage",
        storeAuthStateInCookie: true,
      },
    });
  }

  protected async initialize(): Promise<void> {
    await this._client.initialize();
  }

  public get scopes(): string[] {
    return [`${this.hostname}/Container.Selected`];
  }

  public async login(): Promise<void> {
    await this._client.loginPopup({
      scopes: this.scopes,
      prompt: "select_account",
    });
  }

  public async getToken(): Promise<string> {
    try {
      if (!this._client.getActiveAccount()) {
        throw new InteractionRequiredAuthError("no_account", "No account is signed in");
      }
      const response = await this._client.acquireTokenSilent({
        scopes: this.scopes,
      });
      return response.accessToken;
    } catch (error) {
      if (error instanceof InteractionRequiredAuthError) {
        const response = await this._client.acquireTokenPopup({
          scopes: this.scopes,
        });
        return response.accessToken;
      }
      throw error;
    }
  }
}
