import * as Msal from "@azure/msal-browser";
import * as Constants from "../common/Constants";
import * as Scopes from "../common/Scopes";

export class CustomAppApiAuthProvider {
  private _initialized = false;
  private async initialize(): Promise<void> {
    if (!this._initialized) {
      await this.client.initialize();
      this._initialized = true;
    }
  }
  public readonly client: Msal.PublicClientApplication;
  public static readonly instance: CustomAppApiAuthProvider = new CustomAppApiAuthProvider();

  private constructor() {
    const msalConfig: Msal.Configuration = {
      auth: {
        clientId: Constants.AZURE_CLIENT_ID!,
        authority: Constants.AUTH_AUTHORITY,
      },
      cache: {
        cacheLocation: "localStorage",
        storeAuthStateInCookie: false,
      },
    };
    this.client = new Msal.PublicClientApplication(msalConfig);
  }

  private get redirectUri(): string {
    const protocol = window.location.protocol;
    const hostname = window.location.hostname;
    let port = `:${window.location.port}`;
    if (port === "80" || port === "443") {
      port = "";
    }
    return `${protocol}//${hostname}${port}`;
  }

  public async getToken(scopes: string[] = [Scopes.SAMPLE_API_CONTAINER_MANAGE]): Promise<string> {
    await this.initialize();
    const tokenRequest: Msal.SilentRequest = {
      scopes: scopes,
      prompt: "select_account",
      redirectUri: this.redirectUri,
    };
    const account = this.client.getActiveAccount();
    try {
      if (account) {
        tokenRequest.account = account;
        const result = await this.client.acquireTokenSilent(tokenRequest);
        return result.accessToken;
      }
      throw new Msal.InteractionRequiredAuthError();
    } catch (error) {
      if (error instanceof Msal.InteractionRequiredAuthError) {
        const redirectResponse = await this.client.handleRedirectPromise();
        if (redirectResponse && redirectResponse.account && redirectResponse.accessToken) {
          console.log("Redirect response:", redirectResponse);
          window.location.hash = "";
          this.client.setActiveAccount(redirectResponse.account);
          return redirectResponse.accessToken;
        }
        await this.client.acquireTokenRedirect(tokenRequest);
        return "";
      } else {
        throw error;
      }
    }
  }
}
