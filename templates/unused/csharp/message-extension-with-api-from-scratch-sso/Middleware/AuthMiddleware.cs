using Microsoft.Azure.Functions.Worker.Http;
using Microsoft.Extensions.Logging;

public class AuthMiddleware
{
    private readonly TokenValidator _tokenValidator;
    private readonly string _audience;
    private readonly string _issuer;
    private readonly string[] _allowedTenants;
    private readonly string[] _scopes;

    public AuthMiddleware(TokenValidator tokenValidator, string audience, string issuer, string[] allowedTenants, string[] scopes)
    {
        _tokenValidator = tokenValidator;
        _audience = audience;
        _issuer = issuer;
        _allowedTenants = allowedTenants;
        _scopes = scopes;
    }

    public async Task<bool> ValidateTokenAsync(HttpRequestData req, ILogger logger)
    {
        if (!req.Headers.TryGetValues("Authorization", out var authHeaders))
        {
            return false;
        }

        var token = authHeaders.FirstOrDefault()?.Split(" ").Last();
        if (string.IsNullOrEmpty(token))
        {
            return false;
        }

        try
        {
            var options = new TokenValidationOptions
            {
                AllowedTenants = _allowedTenants,
                Audience = _audience,
                Issuer = _issuer,
                Scopes = _scopes,
            };
            await _tokenValidator.ValidateTokenAsync(token, options);
            return true;
        }
        catch (Exception ex)
        {
            logger.LogError(ex, "Token is invalid");
            return false;
        }
    }
}