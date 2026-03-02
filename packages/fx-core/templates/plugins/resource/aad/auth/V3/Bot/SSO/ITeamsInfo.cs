// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

using Microsoft.Agents.Builder;
using Microsoft.Agents.Extensions.Teams.Models;

namespace {{YOUR_NAMESPACE}}.SSO
{
  /// <summary>
  /// provides utility methods for interactions that occur within Microsoft Teams.
  /// </summary>
  public interface ITeamsInfo
  {
    /// <summary>
    /// Gets the account of a single conversation member. 
    /// This works in one-on-one, group, and teams scoped conversations.
    /// </summary>
    /// <param name="context"> Turn context. </param>
    /// <param name="userId"> ID of the user in question. </param>
    /// <param name="cancellationToken"> cancellation token. </param>
    /// <returns>Team Account Details.</returns>
    Task<TeamsChannelAccount> GetTeamsMemberAsync(ITurnContext context, string userId, CancellationToken cancellationToken = default);
  }
}