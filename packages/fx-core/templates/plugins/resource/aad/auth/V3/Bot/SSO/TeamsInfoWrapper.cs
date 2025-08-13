// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

using Microsoft.Agents.Builder;
using Microsoft.Agents.Extensions.Teams.Connector;
using Microsoft.Agents.Extensions.Teams.Models;

namespace {{YOUR_NAMESPACE}}.SSO
{
  /// <summary>
  /// Helper class used to wrap static method and simplify unit test.
  /// </summary>
  internal class TeamsInfoWrapper : ITeamsInfo
  {
    public Task<TeamsChannelAccount> GetTeamsMemberAsync(ITurnContext context, string userId, CancellationToken cancellationToken = default)
    {
      return TeamsInfo.GetMemberAsync(context, userId, cancellationToken);
    }
  }
}
