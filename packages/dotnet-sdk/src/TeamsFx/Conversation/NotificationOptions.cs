// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

namespace Microsoft.TeamsFx.Conversation
{
    /// <summary>
    /// Options to initialize <see cref="NotificationBot"/>.
    /// </summary>
    [Obsolete("This package will be deprecated by 2026-09. Please use Microsoft 365 Agents SDK (https://www.nuget.org/packages/Microsoft.Agents.Core) instead.")]
    public class NotificationOptions
    {
        /// <summary>
        /// The application ID of the bot.
        /// </summary>
        public string BotAppId { get; set; } = string.Empty;

        /// <summary>
        /// An optional store to persist bot notification target references.
        /// </summary>
        /// <remarks>
        /// <para>
        /// If <c>Store</c> is not provided, a default local store will be used, which stores notification target references into:
        /// </para>
        /// <list type="bullet">
        ///     <item>
        ///         <description>"{$env:TEAMSFX_NOTIFICATION_LOCALSTORE_DIR}/.notification.localstore.json" if running locally.</description>
        ///     </item>
        ///     <item>
        ///         <description>"{$env:TEMP}/.notification.localstore.json" if {$env:RUNNING_ON_AZURE} is set to "1".</description>
        ///     </item>
        ///     <item>
        ///         <description>"{<see cref="Environment.CurrentDirectory"/>}/.notification.localstore.json" if all above environment variables are not set.</description>
        ///     </item>
        /// </list>
        /// <para>
        /// It's recommended to use your own shared store for production environment.
        /// </para>
        /// </remarks>
        public IConversationReferenceStore Store { get; set; }

        /// <summary>
        /// An optional storage to persist bot notification target references.
        /// </summary>
        /// <remarks>
        /// <para>
        /// If <c>Storage</c> is not provided, a default local file storage will be used, which stores notification target references into:
        /// </para>
        /// <list type="bullet">
        ///     <item>
        ///         <description>"{$env:TEAMSFX_NOTIFICATION_LOCALSTORE_DIR}/.notification.localstore.json" if running locally.</description>
        ///     </item>
        ///     <item>
        ///         <description>"{$env:TEMP}/.notification.localstore.json" if {$env:RUNNING_ON_AZURE} is set to "1".</description>
        ///     </item>
        ///     <item>
        ///         <description>"{<see cref="Environment.CurrentDirectory"/>}/.notification.localstore.json" if all above environment variables are not set.</description>
        ///     </item>
        /// </list>
        /// <para>
        /// It's recommended to use your own shared storage for production environment.
        /// </para>
        /// </remarks>
        [Obsolete($"Use {nameof(Store)} to customize the way to persist bot notification target references instead.")]
        public INotificationTargetStorage Storage { get; set; }
    }
}
