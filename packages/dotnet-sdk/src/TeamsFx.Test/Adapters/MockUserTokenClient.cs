// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

using Microsoft.Agents.Connector;
using Microsoft.Agents.Core.Models;

namespace Microsoft.TeamsFx.Test
{
    internal class MockUserTokenClient : IUserTokenClient
    {
        private const string ExceptionExpected = "ExceptionExpected";
        private readonly IDictionary<UserTokenKey, string> _userTokens = new Dictionary<UserTokenKey, string>();
        private readonly IDictionary<ExchangableTokenKey, string> _exchangableToken = new Dictionary<ExchangableTokenKey, string>();
        private readonly IList<TokenMagicCode> _magicCodes = new List<TokenMagicCode>();

        /// <summary>
        /// Adds a fake user token so it can later be retrieved.
        /// </summary>
        /// <param name="connectionName">The connection name.</param>
        /// <param name="channelId">The channel ID.</param>
        /// <param name="userId">The user ID.</param>
        /// <param name="token">The token to store.</param>
        /// <param name="magicCode">The optional magic code to associate with this token.</param>
        public void AddUserToken(string connectionName, string channelId, string userId, string token, string magicCode = null)
        {
            var key = new UserTokenKey()
            {
                ConnectionName = connectionName,
                ChannelId = channelId,
                UserId = userId,
            };

            if (magicCode == null)
            {
                if (_userTokens.ContainsKey(key))
                {
                    _userTokens[key] = token;
                }
                else
                {
                    _userTokens.Add(key, token);
                }
            }
            else
            {
                _magicCodes.Add(new TokenMagicCode()
                {
                    Key = key,
                    MagicCode = magicCode,
                    UserToken = token,
                });
            }
        }

        /// <summary>
        /// Adds a fake exchangeable token so it can be exchanged later.
        /// </summary>
        /// <param name="connectionName">The connection name.</param>
        /// <param name="channelId">The channel ID.</param>
        /// <param name="userId">The user ID.</param>
        /// <param name="exchangableItem">The exchangeable token or resource URI.</param>
        /// <param name="token">The token to store.</param>
        public void AddExchangeableToken(string connectionName, string channelId, string userId, string exchangableItem, string token)
        {
            var key = new ExchangableTokenKey()
            {
                ConnectionName = connectionName,
                ChannelId = channelId,
                UserId = userId,
                ExchangableItem = exchangableItem
            };

            if (_exchangableToken.ContainsKey(key))
            {
                _exchangableToken[key] = token;
            }
            else
            {
                _exchangableToken.Add(key, token);
            }
        }

        /// <summary> Adds an instruction to throw an exception during exchange requests.
        /// </summary>
        /// <param name="connectionName">The connection name.</param>
        /// <param name="channelId">The channel ID.</param>
        /// <param name="userId">The user ID.</param>
        /// <param name="exchangableItem">The exchangeable token or resource URI.</param>
        public void ThrowOnExchangeRequest(string connectionName, string channelId, string userId, string exchangableItem)
        {
            var key = new ExchangableTokenKey()
            {
                ConnectionName = connectionName,
                ChannelId = channelId,
                UserId = userId,
                ExchangableItem = exchangableItem
            };

            if (_exchangableToken.ContainsKey(key))
            {
                _exchangableToken[key] = ExceptionExpected;
            }
            else
            {
                _exchangableToken.Add(key, ExceptionExpected);
            }
        }

        public Task<TokenResponse> GetUserTokenAsync(string userId, string connectionName, string channelId, string magicCode, CancellationToken cancellationToken)
        {
            var key = new UserTokenKey()
            {
                ConnectionName = connectionName,
                ChannelId = channelId, //turnContext.Activity.ChannelId,
                UserId = userId //turnContext.Activity.From.Id,
            };

            if (magicCode != null)
            {
                var magicCodeRecord = _magicCodes.FirstOrDefault(x => key.Equals(x.Key));
                if (magicCodeRecord != null && magicCodeRecord.MagicCode == magicCode)
                {
                    // move the token to long term dictionary
                    AddUserToken(connectionName, key.ChannelId, key.UserId, magicCodeRecord.UserToken);
                    _magicCodes.Remove(magicCodeRecord);
                }
            }

            if (_userTokens.TryGetValue(key, out string token))
            {
                // found
                return Task.FromResult(new TokenResponse()
                {
                    ConnectionName = connectionName,
                    Token = token,
                });
            }

            // not found
            return Task.FromResult<TokenResponse>(null);

        }

        public Task<SignInResource> GetSignInResourceAsync(string connectionName, IActivity activity, string finalRedirect, CancellationToken cancellationToken)
        {
            return Task.FromResult(new SignInResource()
            {
                SignInLink = $"https://fake.com/oauthsignin/{connectionName}/{activity.ChannelId}/{activity?.Recipient?.Id}",
                TokenExchangeResource = new TokenExchangeResource()
                {
                    Id = Guid.NewGuid().ToString(),
                    ProviderId = null,
                    Uri = $"api://{connectionName}/resource"
                }
            });
        }

        public Task SignOutUserAsync(string userId, string connectionName, string channelId, CancellationToken cancellationToken)
        {
            var records = _userTokens.ToArray();
            foreach (var t in records)
            {
                if (t.Key.ChannelId == channelId &&
                    t.Key.UserId == userId &&
                    (connectionName == null || connectionName == t.Key.ConnectionName))
                {
                    _userTokens.Remove(t.Key);
                }
            }

            return Task.CompletedTask;
        }

        public Task<TokenStatus[]> GetTokenStatusAsync(string userId, string channelId, string includeFilter, CancellationToken cancellationToken)
        {
            var filter = includeFilter == null ? null : includeFilter.Split(',');
            var records = _userTokens.
                Where(x =>
                    x.Key.ChannelId == channelId &&
                    x.Key.UserId == userId &&
                    (includeFilter == null || filter.Contains(x.Key.ConnectionName))).
                Select(r => new TokenStatus() { ConnectionName = r.Key.ConnectionName, HasToken = true, ServiceProviderDisplayName = r.Key.ConnectionName }).ToArray();

            if (records.Any())
            {
                return Task.FromResult(records);
            }

            return Task.FromResult<TokenStatus[]>(null);
        }

        public Task<Dictionary<string, TokenResponse>> GetAadTokensAsync(string userId, string connectionName, string[] resourceUrls, string channelId, CancellationToken cancellationToken)
        {
            return Task.FromResult(new Dictionary<string, TokenResponse>());
        }

        public Task<TokenResponse> ExchangeTokenAsync(string userId, string connectionName, string channelId, TokenExchangeRequest exchangeRequest, CancellationToken cancellationToken)
        {
            var exchangableValue = !string.IsNullOrEmpty(exchangeRequest?.Token) ?
                exchangeRequest?.Token :
                exchangeRequest?.Uri;

            var key = new ExchangableTokenKey()
            {
                ChannelId = channelId,
                ConnectionName = connectionName,
                ExchangableItem = exchangableValue,
                UserId = userId,
            };

            if (_exchangableToken.TryGetValue(key, out string token))
            {
                if (token == ExceptionExpected)
                {
                    throw new InvalidOperationException("Exception occurred during exchanging tokens");
                }

                return Task.FromResult(new TokenResponse()
                {
                    ChannelId = key.ChannelId,
                    ConnectionName = key.ConnectionName,
                    Token = token
                });
            }
            else
            {
                return Task.FromResult<TokenResponse>(null);
            }
        }

        public void Dispose()
        {
            GC.SuppressFinalize(this);
        }
    }

    class UserTokenKey
    {
        public string ConnectionName { get; set; }

        public string UserId { get; set; }

        public string ChannelId { get; set; }

        public override bool Equals(object obj)
        {
            var rhs = obj as UserTokenKey;
            if (rhs != null)
            {
                return string.Equals(this.ConnectionName, rhs.ConnectionName, StringComparison.Ordinal) &&
                    string.Equals(this.UserId, rhs.UserId, StringComparison.Ordinal) &&
                    string.Equals(this.ChannelId, rhs.ChannelId, StringComparison.Ordinal);
            }

            return base.Equals(obj);
        }

        public override int GetHashCode()
        {
            return (ConnectionName ?? string.Empty).GetHashCode() +
                (UserId ?? string.Empty).GetHashCode() +
                (ChannelId ?? string.Empty).GetHashCode();
        }
    }

    class ExchangableTokenKey : UserTokenKey
    {
        public string ExchangableItem { get; set; }

        public override bool Equals(object obj)
        {
            var rhs = obj as ExchangableTokenKey;
            if (rhs != null)
            {
                return string.Equals(this.ExchangableItem, rhs.ExchangableItem, StringComparison.Ordinal) &&
                    base.Equals(obj);
            }

            return false;
        }

        public override int GetHashCode()
        {
            return (ExchangableItem ?? string.Empty).GetHashCode() +
                base.GetHashCode();
        }
    }

    class TokenMagicCode
    {
        public UserTokenKey Key { get; set; }

        public string MagicCode { get; set; }

        public string UserToken { get; set; }
    }

}
