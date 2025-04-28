import * as jwt from "jsonwebtoken";

interface KeySet {
  keys: Key[];
}
interface Key {
  kty: string;
  use: string;
  kid: string;
  x5t: string;
  n: string;
  e: string;
  x5c: string[];
}

export class JwtProvider {
  public static readonly KEY_ENDPOINT = "https://login.microsoftonline.com/common/discovery/keys";
  public static readonly AUDIENCE = [
    `${process.env.AZURE_CLIENT_ID}`,
    `api://${process.env.AZURE_CLIENT_ID}`,
  ];

  private static _keys: Map<string, Key> = new Map<string, Key>();

  public constructor(private readonly _token: string) {}

  private static async _loadKeys(): Promise<void> {
    const response = await fetch(JwtProvider.KEY_ENDPOINT);
    const keySet: KeySet = await response.json();
    keySet.keys.forEach((key) => {
      JwtProvider._keys.set(key.kid, key);
    });
  }

  public get token(): string {
    return this._token;
  }

  public get decoded(): jwt.Jwt | null {
    return jwt.decode(this._token, { complete: true });
  }

  public get tid(): string | null {
    const payload = this.decoded?.payload as jwt.JwtPayload;
    if (payload && payload.tid) {
      return payload.tid;
    }
    return null;
  }

  public async getSigningKey(): Promise<string | null> {
    if (JwtProvider._keys.size === 0) {
      await JwtProvider._loadKeys();
    }
    if (this.decoded?.header.kid) {
      const key = JwtProvider._keys.get(this.decoded.header.kid);
      if (key && key.x5c) {
        return `-----BEGIN CERTIFICATE-----\n${key.x5c[0]}\n-----END CERTIFICATE-----`;
      }
    }
    return null;
  }

  public async verify(): Promise<boolean> {
    const decoded = this.decoded;
    if (!this.decoded) {
      return false;
    }

    const signingKey = await this.getSigningKey();
    if (!signingKey) {
      return false;
    }
    const verifyOptions: jwt.VerifyOptions = {
      algorithms: ["RS256"],
      audience: JwtProvider.AUDIENCE,
    };

    try {
      jwt.verify(this._token, signingKey, verifyOptions);
      return true;
    } catch (error) {
      return false;
    }
  }

  public async authorize(scopes: string[] = ["Container.Manage"]): Promise<boolean> {
    if (!this.decoded) {
      return false;
    }

    const payload = this.decoded.payload as jwt.JwtPayload;
    if (!payload || !payload.scp || !payload.scp.split) {
      return false;
    }
    const decodedScopes = payload.scp.split(" ") as string[];
    return (await this.verify()) && scopes.every((scope) => decodedScopes.includes(scope));
  }

  public static fromAuthHeader(authHeader: string | null): JwtProvider | undefined {
    if (!authHeader) {
      return;
    }
    const [bearer, token] = authHeader.split(" ");
    if (!token) {
      return;
    }
    return new JwtProvider(token);
  }

  public static async authorizefromHeader(authHeader: string, scopes: string[]): Promise<boolean> {
    const [bearer, tokenStr] = authHeader.split(" ");
    if (bearer !== "Bearer") {
      return false;
    }
    const token = new JwtProvider(tokenStr);
    return await token.authorize(scopes);
  }
}
