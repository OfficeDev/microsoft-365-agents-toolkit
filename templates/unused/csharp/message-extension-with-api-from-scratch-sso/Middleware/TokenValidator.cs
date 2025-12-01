using System.IdentityModel.Tokens.Jwt;
using Microsoft.IdentityModel.Tokens;
using Newtonsoft.Json.Linq;

public class TokenValidator
{
    private readonly TokenCacheWrapper _cacheWrapper;
    private readonly string _jwksUri;

    public TokenValidator(string tenant = "common", CloudType cloud = CloudType.Public)
    {
        _jwksUri = JwksUriProvider.GetEntraJwksUriAsync(tenant, cloud).GetAwaiter().GetResult();
        _cacheWrapper = new TokenCacheWrapper(GetSigningKeyAsync);
    }

    public async Task ValidateTokenAsync(string token, TokenValidationOptions options)
    {
        var handler = new JwtSecurityTokenHandler();
        var jwtToken = handler.ReadJwtToken(token);

        var key = await _cacheWrapper.GetSigningKeyAsync(jwtToken.Header.Kid);
        var validationParameters = new TokenValidationParameters
        {
            ValidateIssuer = true,
            ValidIssuer = options.Issuer,
            ValidateAudience = true,
            ValidAudience = options.Audience,
            IssuerSigningKey = new JsonWebKey(key)
        };
        handler.ValidateToken(token, validationParameters, out _);
        ValidateScopesAndRoles(jwtToken, options);
        ValidateAllowedTenants(jwtToken, options);
    }

    private async Task<string> GetSigningKeyAsync(string kid)
    {
        using (var httpClient = new HttpClient())
        {
            var jwksResponse = await httpClient.GetStringAsync(_jwksUri);
            var jwks = JObject.Parse(jwksResponse);

            var keys = jwks["keys"];
            if (keys == null)
            {
                throw new Exception("JWKS keys not found.");
            }

            foreach (var key in keys)
            {
                if (key["kid"]?.ToString() == kid)
                {
                    return key.ToString();
                }
            }

            throw new Exception("Key not found.");
        }
    }

    private void ValidateScopesAndRoles(JwtSecurityToken jwt, TokenValidationOptions options)
    {
        if (options.Scopes?.Any() == true || options.Roles?.Any() == true)
        {
            void ValidateClaims(string[] claimsFromTheToken, string[] requiredClaims, string claimsType)
            {
                bool hasAnyRequiredClaim = requiredClaims.Any(claim => claimsFromTheToken.Contains(claim));
                if (!hasAnyRequiredClaim)
                {
                    throw new SecurityTokenException($"JWT does not contain any of the required {claimsType}");
                }
            }

            var scopes = jwt.Claims.FirstOrDefault(c => c.Type == "scp")?.Value?.Split(' ') ?? Array.Empty<string>();
            var roles = jwt.Claims.FirstOrDefault(c => c.Type == "roles")?.Value?.Split(' ') ?? Array.Empty<string>();

            if (options.Scopes != null && options.Roles != null)
            {
                if (scopes.Any())
                {
                    ValidateClaims(scopes, options.Scopes, "scopes");
                }
                else if (roles.Any())
                {
                    ValidateClaims(roles, options.Roles, "roles");
                }
            }
            else if (options.Scopes != null)
            {
                ValidateClaims(scopes, options.Scopes, "scopes");
            }
            else if (options.Roles != null)
            {
                ValidateClaims(roles, options.Roles, "roles");
            }
        }
    }

    private void ValidateAllowedTenants(JwtSecurityToken jwt, TokenValidationOptions options)
    {
        if (options.AllowedTenants?.Any() == true)
        {
            var tenantId = jwt.Claims.FirstOrDefault(c => c.Type == "tid")?.Value;

            if (tenantId == null || !options.AllowedTenants.Contains(tenantId))
            {
                throw new SecurityTokenException($"JWT tid is not allowed. Allowed tenants: {string.Join(", ", options.AllowedTenants)}");
            }
        }
    }
}

public class TokenValidationOptions
{
    public string[] AllowedTenants { get; set; }
    public string Audience { get; set; }
    public string Issuer { get; set; }
    public string[] Scopes { get; set; }
    public string[] Roles { get; set; }
}