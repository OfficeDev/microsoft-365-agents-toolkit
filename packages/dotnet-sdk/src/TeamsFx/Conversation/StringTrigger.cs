// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

namespace Microsoft.TeamsFx.Conversation
{
    /// <summary>
    /// Represents a command trigger that triggered by string match.
    /// </summary>
    ///  <seealso cref="ITeamsCommandHandler"/>
    [Obsolete("This package will be deprecated by 2026-09. Please use Microsoft 365 Agents SDK (https://www.nuget.org/packages/Microsoft.Agents.Hosting.AspNetCore) instead.")]
    public class StringTrigger : ITriggerPattern
    {
        /// <summary>
        /// Gets or sets the string pattern used to match the input.
        /// </summary>
        public string Pattern { get; set; }

        /// <summary>
        /// Initializes a new instance of the <see cref="StringTrigger"/> class.
        /// </summary>
        /// <param name="pattern">The string used to compare with the input.</param>
        /// <exception cref="ArgumentNullException"><paramref name="pattern"/>is null.</exception>
        public StringTrigger(string pattern)
        { 
            Pattern = pattern ?? throw new ArgumentNullException(nameof(pattern));
        }

        /// <inheritdoc/>
        public bool ShouldTrigger(string input)
        {
            return string.Equals(input, this.Pattern, StringComparison.OrdinalIgnoreCase);
        }
    }
}
