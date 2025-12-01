using System.ComponentModel;
using System.Text.Json.Serialization;

namespace {{SafeProjectName}}.Interop.TeamsSDK;

public class TeamsContext
{
    public TeamsMeetingContext Meeting { get; set; }

    public TeamsAppContext App { get; set; }

    public TeamsChannelContext Channel { get; set; }

    public TeamsChatContext Chat { get; set; }

    public TeamsPageContext Page { get; set; }

    public TeamsTeamContext Team { get; set; }

    public TeamsUserContext User { get; set; }

    public TeamsSharepointSiteContext SharepointSite { get; set; }
}

public class TeamsMeetingContext
{
    /// <summary>
    /// Meeting Id used by tab when running in meeting context.
    /// </summary>
    public string Id { get; set; }
}

public class TeamsAppContext
{
    /// <summary>
    /// Personal app icon y coordinate position.
    /// </summary>
    public double IconPositionVertical { get; set; }

    /// <summary>
    /// Unique ID for the current session for use in correlating telemetry data.
    /// </summary>
    public string SessionId { get; set; }

    public TeamsHostContext Host { get; set; }

    /// <summary>
    /// The current locale that the user has configured for the app formatted as languageId-countryId (for example, en-us).
    /// </summary>
    public string Locale { get; set; }

    /// <summary>
    /// More detailed locale info from the user's OS if available. Can be used together with the @microsoft/globe NPM
    /// package to ensure your app respects the user's OS date and time format configuration.
    /// </summary>
    public LocaleInfo OSLocaleInfo { get; set; }

    /// <summary>
    /// The ID of the parent message from which this task module was launched. This is only available in task modules launched from bot cards.
    /// </summary>
    public string ParentMessageId { get; set; }

    /// <summary>
    /// Time when the user clicked on the tab.
    /// </summary>
    public double? UserClickTime { get; set; }
}

public class TeamsChannelContext
{
    /// <summary>
    /// The Microsoft Teams ID for the channel with which the content is associated.
    /// </summary>
    public string Id { get; set; }

    /// <summary>
    /// The name for the channel with which the content is associated.
    /// </summary>
    public string DisplayName { get; set; }

    /// <summary>
    /// The relative path to the SharePoint folder associated with the channel.
    /// </summary>
    public string RelativeUrl { get; set; }

    /// <summary>
    /// The type of the channel with which the content is associated.
    /// </summary>
    [JsonConverter(typeof(EnumDescriptionConverter<ChannelType>))]
    public ChannelType MembershipType { get; set; }

    /// <summary>
    /// The OneNote section ID that is linked to the channel.
    /// </summary>
    public string DefaultOneNoteSectionId { get; set; }
}

public class TeamsChatContext
{
    /// <summary>
    /// The Microsoft Teams ID for the chat with which the content is associated.
    /// </summary>
    public string Id { get; set; }
}

public class TeamsPageContext
{
    /// <summary>
    /// The developer-defined unique ID for the entity this content points to.
    /// </summary>
    public string Id { get; set; }

    /// <summary>
    /// The context where tab url is loaded (content, task, setting, remove, sidePanel).
    /// </summary>
    [JsonConverter(typeof(EnumDescriptionConverter<FrameContext>))]
    public FrameContext FrameContext { get; set; }

    /// <summary>
    /// Indication whether the tab is in full-screen mode.
    /// </summary>
    public bool? IsFullScreen { get; set; }

    /// <summary>
    /// Indication whether the tab is in a pop out window.
    /// </summary>
    public bool? IsMultiWindow { get; set; }

    /// <summary>
    /// The developer-defined unique ID for the sub-entity this content points to. This field should be used to restore to a
    /// specific state within an entity, such as scrolling to or activating a specific piece of content.
    /// </summary>
    public string SubPageId { get; set; }
}

public class TeamsTeamContext
{
    /// <summary>
    /// The Office 365 group ID for the team with which the content is associated. This field is available only when the identity permission is requested in the manifest.
    /// </summary>
    public string GroupId { get; set; }

    /// <summary>
    /// Indicates whether team is archived. Apps should use this as a signal to prevent any changes to content associated with archived teams.
    /// </summary>
    public bool? IsArchived { get; set; }

    /// <summary>
    /// The Microsoft Teams ID for the team with which the content is associated.
    /// </summary>
    public string InternalId { get; set; }

    /// <summary>
    /// Indicates the team type, currently used to distinguish between different team types in Office 365 for Education.
    /// </summary>
    public TeamType? Type { get; set; }

    /// <summary>
    /// The user's role in the team. Because a malicious party can run your content in a browser,
    /// this value should be used only as a hint as to the user's role, and never as proof of her role.
    /// </summary>
    public UserTeamRole? UserRole { get; set; }
}

public class TeamsUserContext
{
    /// <summary>
    /// Represents whether calling is allowed for the current logged in User.
    /// </summary>
    public bool? IsCallingAllowed { get; set; }

    /// <summary>
    /// Represents whether PSTN calling is allowed for the current logged in User.
    /// </summary>
    public bool? IsPSTNCallingAllowed { get; set; }

    /// <summary>
    /// A value suitable for use as a login_hint when authenticating with Microsoft Entra. Because a malicious party can run your
    /// content in a browser, this value should be used only as a hint as to who the user is and never as proof of identity.
    /// This field is available only when the identity permission is requested in the manifest.
    /// </summary>
    public string LoginHint { get; set; }

    public TeamsUserTenantContext Tenant { get; set; }

    /// <summary>
    /// The license type for the current user.
    /// </summary>
    public string LicenseType { get; set; }

    /// <summary>
    /// The Microsoft Entra object id of the current user. Because a malicious party run your content in a browser,
    /// this value should be used only as a hint as to who the user is and never as proof of identity. This field
    /// is available only when the identity permission is requested in the manifest.
    /// </summary>
    public string Id { get; set; }

    /// <summary>
    /// The UPN of the current user. This may be an externally-authenticated UPN (e.g., guest users). Because a
    /// malicious party run your content in a browser, this value should be used only as a hint as to who the
    /// user is and never as proof of identity. This field is available only when the identity permission is
    /// requested in the manifest.
    /// </summary>
    public string UserPrincipalName { get; set; }
}

public class TeamsSharepointSiteContext
{
    /// <summary>
    /// The domain of the root SharePoint site associated with the team.
    /// </summary>
    public string Domain { get; set; }

    /// <summary>
    /// The relative path to the SharePoint site associated with the team.
    /// </summary>
    public string Path { get; set; }

    /// <summary>
    /// The root SharePoint site associated with the team.
    /// </summary>
    public string Url { get; set; }
}

public class TeamsHostContext
{
    /// <summary>
    /// The type of the host client. Possible values are : android, ios, web, desktop, rigel.
    /// </summary>
    [JsonConverter(typeof(EnumDescriptionConverter<HostClientType>))]
    public HostClientType ClientType { get; set; }

    /// <summary>
    /// Current ring ID.
    /// </summary>
    public string RingId { get; set; }

    /// <summary>
    /// Unique ID for the current Teams session for use in correlating telemetry data.
    /// </summary>
    public string SessionId { get; set; }
}

public class TeamsUserTenantContext
{
    /// <summary>
    /// The type of license for the current users tenant.
    /// </summary>
    public string TeamsSKU { get; set; }

    /// <summary>
    /// The Microsoft Entra tenant ID of the current user. Because a malicious party can run your content in a browser,
    /// this value should be used only as a hint as to who the user is and never as proof of identity. This field
    /// is available only when the identity permission is requested in the manifest.
    /// </summary>
    public string Id { get; set; }
}

public enum ChannelType
{
    [Description("Private")] Private,
    [Description("Regular")] Regular
}

public enum FrameContext
{
    [Description("authentication")] Authentication,
    [Description("content")] Content,
    [Description("remove")] Remove,
    [Description("settings")] Settings,
    [Description("sidePanel")] SidePanel,
    [Description("stage")] Stage,
    [Description("task")] Task
}

public enum HostClientType
{
    [Description("android")] Android,
    [Description("desktop")] Desktop,
    [Description("ios")] iOS,
    [Description("rigel")] Rigel,
    [Description("web")] Web
}

public enum Platform
{
    [Description("windows")] Windows,
    [Description("macos")] macOS
}

public enum TeamType
{
    Standard = 0,
    Edu = 1,
    Class = 2,
    Plc = 3,
    Staff = 4
}

public enum UserTeamRole
{
    Admin = 0,
    User = 1,
    Guest = 2
}

public class LocaleInfo
{
    public string LongDate { get; set; }
    public string LongTime { get; set; }
    [JsonConverter(typeof(EnumDescriptionConverter<Platform>))]
    public Platform Platform { get; set; }
    public string RegionalFormat { get; set; }
    public string ShortDate { get; set; }
    public string ShortTime { get; set; }
}