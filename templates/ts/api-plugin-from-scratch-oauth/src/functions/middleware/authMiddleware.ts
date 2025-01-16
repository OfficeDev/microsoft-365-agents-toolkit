import { HttpRequest } from "@azure/functions";
import { TokenValidator, EntraJwtPayload } from "./tokenValidator";
import config from "./config";
import { getEntraJwksUri, CloudType } from "./utils";

// Export symbols app devs will need to use
export { CloudType } from "./utils";
export { EntraJwtPayload } from "./tokenValidator";

/**
 * Middleware function to handle authorization using JWT.
 *
 * @param {HttpRequest} req - The HTTP request.
 * @returns {Promise<EntraJwtPayload | false>} - A promise that resolves to an array of JWT claims or false if authentication failed
 */
export async function authMiddleware(req: HttpRequest,
                                     scope: string | [string],
                                     allowedTenants: [string] = [config.aadAppTenantId],
                                     cloud: CloudType = CloudType.Public,
                                     issuer: string = `https://login.microsoftonline.com/${config.aadAppTenantId}/v2.0`
                                    ): Promise<EntraJwtPayload | false> {

  // Get the token from the request headers
  const token = req.headers.get("authorization")?.split(" ")[1];
  if (!token) {
    return false;
  }

  try {
    // Get the JWKS URL for the specified Microsoft Entra cloud
    const entraJwksUri = await getEntraJwksUri(config.aadAppTenantId, cloud);

    // Create a new token validator with the JWKS URL
    const validator = new TokenValidator({
      jwksUri: entraJwksUri,
    });

    const options = {
      allowedTenants: allowedTenants,
      audience: config.aadAppClientId,
      issuer: issuer,
      scp: typeof scope === 'string' ? [scope] : scope
    };
    // Validate the token
    const claims = await validator.validateToken(token, options);

    return claims;

  } catch (err) {

    // Handle JWT verification errors
    console.error("Token is invalid:", err);
    return false;

  }
}
