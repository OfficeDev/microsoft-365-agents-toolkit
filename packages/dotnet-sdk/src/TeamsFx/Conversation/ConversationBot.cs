namespace Microsoft.TeamsFx.Conversation
{
    using Microsoft.Agents.Hosting.AspNetCore;

    /// <summary>
    /// Provide utilities for bot conversation, including:
    /// <list type="bullet">
    ///     <item>
    ///         <description>Send notification to varies targets (e.g., member, group, channel).</description>
    ///     </item>
    ///     <item>
    ///         <description>Handle command and send a response to Teams.</description>
    ///     </item>
    /// </list>
    /// </summary>
    /// <remarks>
    /// <para>
    /// Set <c>Adapter</c> in <see cref="ConversationOptions"/> to use your own bot adapter.
    /// </para>
    /// <para>
    /// For notification, set <c>Notification.Storage</c> in <see cref="ConversationOptions"/> to use your own storage implementation.
    /// </para>
    /// <para>
    /// For command, set <c>Command.Commands</c> in <see cref="CommandOptions"/> to register your command handlers.
    /// </para>
    /// </remarks>
    public class ConversationBot
    {
        /// <summary>
        /// The bot adapter.
        /// </summary>
        public CloudAdapter Adapter { get; private set; }

        /// <summary>
        /// The entry point of notification.
        /// </summary>
        public NotificationBot Notification { get; private set; }

        /// <summary>
        /// The entry point of command.
        /// </summary>
        public CommandBot Command { get; private set; }

        /// <summary>
        /// The entry point of adaptive card action.
        /// </summary>
        public CardActionBot CardAction { get; private set; }

        /// <summary>
        /// Creates new instance of the <see cref="ConversationBot"/>.
        /// </summary>
        /// <param name="options">Initialize options.</param>
        /// <exception cref="ArgumentNullException">Throws if provided parameter is null.</exception>
        /// <exception cref="ArgumentException">Throws if provided parameter is invalid.</exception>
        /// <remarks>
        /// It's recommended to create your own adapter and storage for production environment.
        /// </remarks>
        public ConversationBot(ConversationOptions options)
        {
            if (options == null)
            {
                throw new ArgumentNullException(nameof(options));
            }

            Adapter = options.Adapter ?? throw new ArgumentException("Adapter is null.", nameof(options));

            if (options.Notification != null)
            {
                Notification = new NotificationBot(Adapter, options.Notification);
            }

            if (options.Command != null)
            {
                Command = new CommandBot(Adapter, options.Command);
            }

            if (options.CardAction != null)
            {
                CardAction = new CardActionBot(Adapter, options.CardAction);
            }
        }
    }
}
