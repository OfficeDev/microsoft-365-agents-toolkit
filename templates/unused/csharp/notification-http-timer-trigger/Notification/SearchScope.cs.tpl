namespace {{SafeProjectName}}.Notification
{
    [Flags]
    public enum SearchScope
    {
        /// <summary>
        /// Search members from the installations in personal chat only.
        /// </summary>
        Person = 1,

        /// <summary>
        /// Search members from the installations in group chat only.
        /// </summary>
        Group = 2,

        /// <summary>
        /// Search members from the installations in Teams channel only.
        /// </summary>
        Channel = 4,

        /// <summary>
        /// Search members from all installations including personal chat, group chat and Teams channel.
        /// </summary>
        All = Person | Group | Channel
    }
}
