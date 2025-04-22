// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

namespace Microsoft.TeamsFx.Conversation
{
    using Microsoft.Agents.Builder;
    using Microsoft.Agents.Core.Models;
    using Microsoft.Agents.Extensions.Teams.Models;

    static internal class ITurnContextExtensions
    {
        static internal string GetTeamsBotInstallationId(this ITurnContext context)
        {
            string result = null;
            var activity = context?.Activity;
            if (activity != null)
            {
                var channelData = activity.GetChannelData<TeamsChannelData>();
                if (channelData != null)
                {
                    result = channelData?.Team?.Id;
                }

                // Fallback to use conversation id.
                // The conversation id is equal to team id only when the bot app is installed into the General channel.
                if (result == null && activity.Conversation?.Name == null)
                {
                    result = activity.Conversation?.Id;
                }
            }

            return result;
        }
    }
}
