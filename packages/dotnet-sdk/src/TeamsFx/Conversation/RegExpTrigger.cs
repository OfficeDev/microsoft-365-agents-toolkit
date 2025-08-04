// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

namespace Microsoft.TeamsFx.Conversation
{
    using System.Text.RegularExpressions;
    
    /// <summary>
    /// Represents a command trigger that triggered by regular expression match.
    /// </summary>
    /// <seealso cref="ITeamsCommandHandler"/>
    [Obsolete("This package will be deprecated by 2026-09. Please use Microsoft 365 Agents SDK (https://www.nuget.org/packages/Microsoft.Agents.Core) instead.")]
    public class RegExpTrigger : ITriggerPattern
    {
        /// <summary>
        /// Gets or sets the regular expression used to match the input.
        /// </summary>
        public Regex Pattern { get; set; }

        /// <summary>
        /// Initializes a new instance of the <see cref="RegExpTrigger"/> class.
        /// </summary>
        /// <param name="pattern">The regular expression pattern string.</param>
        /// <exception cref="ArgumentNullException"><paramref name="pattern"/> is null.</exception>
        public RegExpTrigger(string pattern)
        {
            if (string.IsNullOrEmpty(pattern))
            {
                throw new ArgumentNullException(nameof(pattern));
            }

            Pattern = new Regex(pattern, RegexOptions.IgnoreCase);
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="RegExpTrigger"/> class.
        /// </summary>
        /// <param name="pattern">The regular expression pattern.</param>
        /// <exception cref="ArgumentNullException"><paramref name="pattern"/> is null.</exception>
        public RegExpTrigger(Regex pattern)
        {
            Pattern = pattern ?? throw new ArgumentNullException(nameof(pattern));
        }

        /// <inheritdoc/>
        public bool ShouldTrigger(string input)
        {
            if (string.IsNullOrEmpty(input))
            {
                return false;
            }

            return Pattern.IsMatch(input);
        }
    }
}
