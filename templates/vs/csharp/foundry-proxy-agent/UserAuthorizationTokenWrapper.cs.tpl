using Azure.Core;
using Microsoft.Agents.Builder;
using Microsoft.Agents.Builder.App.UserAuth;
using System.IdentityModel.Tokens.Jwt;

namespace {{SafeProjectName}}
{
    /// <summary>
    /// This class wraps the UserAuthorization to provide a TokenCredential implementation as the AI Foundry agent expects a TokenCredential to be used for authentication.
    /// Note to be able to authenticate with the AI Foundry agent, the application that was used to create the user JWT token must have the 'Azure Machine Learning Services' => 'user_impersonation' scope configured in the Azure portal.
    /// </summary>
    public class UserAuthorizationTokenWrapper : TokenCredential
    {
        private readonly UserAuthorization _userAuthorization;
        private readonly string _handlerName;
        private readonly ITurnContext _turnContext;
        public UserAuthorizationTokenWrapper(UserAuthorization userAuthorization, ITurnContext turnContext, string handlerName)
        {
            _userAuthorization = userAuthorization;
            _handlerName = handlerName ?? throw new ArgumentNullException(nameof(handlerName));
            _turnContext = turnContext ?? throw new ArgumentNullException(nameof(turnContext));
        }

        public override AccessToken GetToken(TokenRequestContext requestContext, CancellationToken cancellationToken)
        {
#pragma warning disable CA2012
            return GetTokenAsync(requestContext, cancellationToken).Result;
#pragma warning restore CA2012
        }

        /// <summary>
        /// This method exchanges the current user's turn token for a JWT token that can be used to authenticate with the AI Foundry agent.
        /// </summary>
        /// <param name="requestContext"></param>
        /// <param name="cancellationToken"></param>
        /// <returns></returns>
        /// <exception cref="InvalidOperationException"></exception>
        public override async ValueTask<AccessToken> GetTokenAsync(TokenRequestContext requestContext, CancellationToken cancellationToken)
        {


            // Get the JWT token for SSO using the UserAuthorization service. No Token exchange is needed as the azure bot service Oauth Connection is doing this for us.
            var jwtToken = await _userAuthorization.GetTurnTokenAsync(_turnContext, handlerName: _handlerName, cancellationToken: cancellationToken).ConfigureAwait(false);

            // Convert the JWT token to a Azure AccessToken.
            var handler = new JwtSecurityTokenHandler();
            var jwt = handler.ReadJwtToken(jwtToken);
            long? expClaim = jwt.Payload.Expiration;
            if (expClaim == null)
                throw new InvalidOperationException("JWT does not contain an 'exp' claim.");
            var expiresOn = DateTimeOffset.FromUnixTimeSeconds((long)expClaim);

            return new AccessToken(jwtToken, expiresOn);
        }

    }
}
