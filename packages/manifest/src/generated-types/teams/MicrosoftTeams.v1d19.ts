// To parse this data:
//
//   import { Convert, MicrosoftTeamsV1D19 } from "./file";
//
//   const microsoftTeamsV1D19 = Convert.toMicrosoftTeamsV1D19(json);
//
// These functions will throw an error if the JSON doesn't
// match the expected interface, even if the JSON is valid.

export interface MicrosoftTeamsV1D19 {
    $schema?: string;
    /**
     * A color to use in conjunction with the icon. The value must be a valid HTML color code
     * starting with '#', for example `#4464ee`.
     */
    accentColor: string;
    activities?: Activities;
    /**
     * Specify and consolidates authorization related information for the App.
     */
    authorization?: MicrosoftTeamsV1D19Authorization;
    /**
     * The set of bots for this app. Currently only one bot per app is supported.
     */
    bots?: Bot[];
    /**
     * The set of compose extensions for this app. Currently only one compose extension per app
     * is supported.
     */
    composeExtensions?: ComposeExtension[];
    /**
     * A list of tenant configured properties for an app
     */
    configurableProperties?: ConfigurableProperty[];
    /**
     * These are tabs users can optionally add to their channels and 1:1 or group chats and
     * require extra configuration before they are added. Configurable tabs are not supported in
     * the personal scope. Currently only one configurable tab per app is supported.
     */
    configurableTabs?: ConfigurableTab[];
    /**
     * The set of Office365 connectors for this app. Currently only one connector per app is
     * supported.
     */
    connectors?:    Connector[];
    copilotAgents?: CopilotAgents;
    /**
     * Defines the list of cards which could be pinned to dashboards that can provide summarized
     * view of information relevant to user.
     */
    dashboardCards?: DashboardCard[];
    /**
     * A value indicating whether an app is blocked by default until admin allows it
     */
    defaultBlockUntilAdminAction?: boolean;
    /**
     * When a group install scope is selected, this will define the default capability when the
     * user installs the app
     */
    defaultGroupCapability?: DefaultGroupCapability;
    /**
     * The install scope defined for this app by default. This will be the option displayed on
     * the button when a user tries to add the app
     */
    defaultInstallScope?: DefaultInstallScope;
    description:          Description;
    developer:            Developer;
    /**
     * Specify the native features on a user's device that your app may request access to.
     */
    devicePermissions?: DevicePermission[];
    extensions?:        ElementExtension[];
    /**
     * Specify the app's Graph connector configuration. If this is present then
     * webApplicationInfo.id must also be specified.
     */
    graphConnector?: GraphConnector;
    icons:           Icons;
    /**
     * A unique identifier for this app. This id must be a GUID.
     */
    id: string;
    /**
     * A value indicating whether a personal app is rendered without a tab header-bar
     */
    isFullScreen?:     boolean;
    localizationInfo?: LocalizationInfo;
    /**
     * The version of the schema this manifest is using. This schema version supports extending
     * Teams apps to other parts of the Microsoft 365 ecosystem. More info at
     * https://aka.ms/extendteamsapps.
     */
    manifestVersion: ManifestVersion;
    /**
     * Specify meeting extension definition.
     */
    meetingExtensionDefinition?: MeetingExtensionDefinition;
    name:                        Name;
    /**
     * Specifies the permissions the app requests from users.
     */
    permissions?: Permission[];
    /**
     * The url to the page that provides additional app information for the admins
     */
    publisherDocsUrl?: string;
    /**
     * A value indicating whether or not show loading indicator when app/tab is loading
     */
    showLoadingIndicator?: boolean;
    /**
     * A set of tabs that may be 'pinned' by default, without the user adding them manually.
     * Static tabs declared in personal scope are always pinned to the app's personal
     * experience. Static tabs do not currently support the 'teams' scope.
     */
    staticTabs?: StaticTab[];
    /**
     * Subscription offer associated with this app.
     */
    subscriptionOffer?: SubscriptionOffer;
    /**
     * List of 'non-standard' channel types that the app supports. Note: Channels of standard
     * type are supported by default if the app supports team scope.
     */
    supportedChannelTypes?: SupportedChannelType[];
    /**
     * A list of valid domains from which the tabs expect to load any content. Domain listings
     * can include wildcards, for example `*.example.com`. If your tab configuration or content
     * UI needs to navigate to any other domain besides the one use for tab configuration, that
     * domain must be specified here.
     */
    validDomains?: string[];
    /**
     * The version of the app. Changes to your manifest should cause a version change. This
     * version string must follow the semver standard (http://semver.org).
     */
    version: string;
    /**
     * Specify your AAD App ID and Graph information to help users seamlessly sign into your AAD
     * app.
     */
    webApplicationInfo?: WebApplicationInfo;
}

export interface Activities {
    /**
     * Specify the types of activites that your app can post to a users activity feed
     */
    activityTypes?: ActivityType[];
}

export interface ActivityType {
    description:  string;
    templateText: string;
    type:         string;
}

/**
 * Specify and consolidates authorization related information for the App.
 */
export interface MicrosoftTeamsV1D19Authorization {
    /**
     * List of permissions that the app needs to function.
     */
    permissions?: Permissions;
}

/**
 * List of permissions that the app needs to function.
 */
export interface Permissions {
    /**
     * Permissions that must be granted on a per resource instance basis.
     */
    resourceSpecific?: ResourceSpecific[];
}

export interface ResourceSpecific {
    /**
     * The name of the resource-specific permission.
     */
    name: string;
    /**
     * The type of the resource-specific permission: delegated vs application.
     */
    type: ResourceSpecificType;
}

/**
 * The type of the resource-specific permission: delegated vs application.
 */
export enum ResourceSpecificType {
    Application = "Application",
    Delegated = "Delegated",
}

export interface Bot {
    /**
     * The Microsoft App ID specified for the bot in the Bot Framework portal
     * (https://dev.botframework.com/bots)
     */
    botId: string;
    /**
     * The list of commands that the bot supplies, including their usage, description, and the
     * scope for which the commands are valid. A separate command list should be used for each
     * scope.
     */
    commandLists?:  CommandList[];
    configuration?: Configuration;
    /**
     * A value indicating whether or not the bot is a one-way notification only bot, as opposed
     * to a conversational bot.
     */
    isNotificationOnly?: boolean;
    /**
     * This value describes whether or not the bot utilizes a user hint to add the bot to a
     * specific channel.
     */
    needsChannelSelector?: boolean;
    /**
     * Specifies whether the bot offers an experience in the context of a channel in a team, in
     * a 1:1 or group chat, or in an experience scoped to an individual user alone. These
     * options are non-exclusive.
     */
    scopes: CommandListScope[];
    /**
     * A value indicating whether the bot supports audio calling.
     */
    supportsCalling?: boolean;
    /**
     * A value indicating whether the bot supports uploading/downloading of files.
     */
    supportsFiles?: boolean;
    /**
     * A value indicating whether the bot supports video calling.
     */
    supportsVideo?: boolean;
}

export interface CommandList {
    commands: CommandListCommand[];
    /**
     * Specifies the scopes for which the command list is valid
     */
    scopes: CommandListScope[];
}

export interface CommandListCommand {
    /**
     * A simple text description or an example of the command syntax and its arguments.
     */
    description: string;
    /**
     * The bot command name
     */
    title: string;
}

export enum CommandListScope {
    GroupChat = "groupChat",
    Personal = "personal",
    Team = "team",
}

export interface Configuration {
    groupChat?: GroupChat;
    team?:      GroupChat;
}

export interface GroupChat {
    fetchTask?: boolean;
    taskInfo?:  TaskInfo;
}

export interface TaskInfo {
    /**
     * Dialog height - either a number in pixels or default layout such as 'large', 'medium', or
     * 'small'
     */
    height?: string;
    /**
     * Initial dialog title
     */
    title?: string;
    /**
     * Initial webview URL
     */
    url?: string;
    /**
     * Dialog width - either a number in pixels or default layout such as 'large', 'medium', or
     * 'small'
     */
    width?: string;
}

export interface ComposeExtension {
    /**
     * A relative file path to the api specification file in the manifest package.
     */
    apiSpecificationFile?: string;
    /**
     * Object capturing authorization information.
     */
    authorization?: ComposeExtensionAuthorization;
    /**
     * The Microsoft App ID specified for the bot powering the compose extension in the Bot
     * Framework portal (https://dev.botframework.com/bots)
     */
    botId?: string;
    /**
     * A value indicating whether the configuration of a compose extension can be updated by the
     * user.
     */
    canUpdateConfiguration?: boolean | null;
    commands?:               ComposeExtensionCommand[];
    /**
     * Type of the compose extension.
     */
    composeExtensionType?: ComposeExtensionType;
    /**
     * A list of handlers that allow apps to be invoked when certain conditions are met
     */
    messageHandlers?: MessageHandler[];
}

/**
 * Object capturing authorization information.
 */
export interface ComposeExtensionAuthorization {
    /**
     * Object capturing details needed to do service auth. It will be only present when auth
     * type is apiSecretServiceAuth.
     */
    apiSecretServiceAuthConfiguration?: APISecretServiceAuthConfiguration;
    /**
     * Enum of possible authentication types.
     */
    authType?: AuthType;
    /**
     * Object capturing details needed to do single aad auth flow. It will be only present when
     * auth type is entraId.
     */
    microsoftEntraConfiguration?: MicrosoftEntraConfiguration;
}

/**
 * Object capturing details needed to do service auth. It will be only present when auth
 * type is apiSecretServiceAuth.
 */
export interface APISecretServiceAuthConfiguration {
    /**
     * Registration id returned when developer submits the api key through Developer Portal.
     */
    apiSecretRegistrationId?: string;
}

/**
 * Enum of possible authentication types.
 */
export enum AuthType {
    APISecretServiceAuth = "apiSecretServiceAuth",
    MicrosoftEntra = "microsoftEntra",
    None = "none",
}

/**
 * Object capturing details needed to do single aad auth flow. It will be only present when
 * auth type is entraId.
 */
export interface MicrosoftEntraConfiguration {
    /**
     * Boolean indicating whether single sign on is configured for the app.
     */
    supportsSingleSignOn?: boolean;
}

export interface ComposeExtensionCommand {
    /**
     * A relative file path for api response rendering template file.
     */
    apiResponseRenderingTemplateFile?: string;
    /**
     * Context where the command would apply
     */
    context?: CommandContext[];
    /**
     * Description of the command.
     */
    description?: string;
    /**
     * A boolean value that indicates if it should fetch task module dynamically
     */
    fetchTask?: boolean;
    /**
     * Id of the command.
     */
    id: string;
    /**
     * A boolean value that indicates if the command should be run once initially with no
     * parameter.
     */
    initialRun?:    boolean;
    parameters?:    Parameter[];
    samplePrompts?: SamplePrompt[];
    /**
     * Semantic description for the command.
     */
    semanticDescription?: string;
    taskInfo?:            TaskInfo;
    /**
     * Title of the command.
     */
    title: string;
    /**
     * Type of the command
     */
    type?: CommandType;
}

export enum CommandContext {
    CommandBox = "commandBox",
    Compose = "compose",
    Message = "message",
}

export interface Parameter {
    /**
     * The choice options for the parameter
     */
    choices?: Choice[];
    /**
     * Description of the parameter.
     */
    description?: string;
    /**
     * Type of the parameter
     */
    inputType?: InputType;
    /**
     * The value indicates if this parameter is a required field.
     */
    isRequired?: boolean;
    /**
     * Name of the parameter.
     */
    name: string;
    /**
     * Semantic description for the parameter.
     */
    semanticDescription?: string;
    /**
     * Title of the parameter.
     */
    title: string;
    /**
     * Initial value for the parameter
     */
    value?: string;
}

export interface Choice {
    /**
     * Title of the choice
     */
    title: string;
    /**
     * Value of the choice
     */
    value: string;
}

/**
 * Type of the parameter
 */
export enum InputType {
    Choiceset = "choiceset",
    Date = "date",
    Number = "number",
    Text = "text",
    Textarea = "textarea",
    Time = "time",
    Toggle = "toggle",
}

export interface SamplePrompt {
    /**
     * This string will hold the sample prompt
     */
    text: string;
}

/**
 * Type of the command
 */
export enum CommandType {
    Action = "action",
    Query = "query",
}

/**
 * Type of the compose extension.
 */
export enum ComposeExtensionType {
    APIBased = "apiBased",
    BotBased = "botBased",
}

export interface MessageHandler {
    /**
     * Type of the message handler
     */
    type:  MessageHandlerType;
    value: Value;
}

/**
 * Type of the message handler
 */
export enum MessageHandlerType {
    Link = "link",
}

export interface Value {
    /**
     * A list of domains that the link message handler can register for, and when they are
     * matched the app will be invoked
     */
    domains?: string[];
    /**
     * A boolean that indicates whether the app's link message handler supports anonymous invoke
     * flow.
     */
    supportsAnonymizedPayloads?: boolean;
    [property: string]: any;
}

export enum ConfigurableProperty {
    AccentColor = "accentColor",
    DeveloperURL = "developerUrl",
    LargeImageURL = "largeImageUrl",
    LongDescription = "longDescription",
    Name = "name",
    PrivacyURL = "privacyUrl",
    ShortDescription = "shortDescription",
    SmallImageURL = "smallImageUrl",
    TermsOfUseURL = "termsOfUseUrl",
}

export interface ConfigurableTab {
    /**
     * A value indicating whether an instance of the tab's configuration can be updated by the
     * user after creation.
     */
    canUpdateConfiguration?: boolean;
    /**
     * The url to use when configuring the tab.
     */
    configurationUrl: string;
    /**
     * The set of contextItem scopes that a tab belong to
     */
    context?: ConfigurableTabContext[];
    /**
     * The set of meetingSurfaceItem scopes that a tab belong to
     */
    meetingSurfaces?: MeetingSurface[];
    /**
     * Specifies whether the tab offers an experience in the context of a channel in a team, in
     * a 1:1 or group chat, or in an experience scoped to an individual user alone. These
     * options are non-exclusive. Currently, configurable tabs are only supported in the teams
     * and groupchats scopes.
     */
    scopes: ConfigurableTabScope[];
    /**
     * A relative file path to a tab preview image for use in SharePoint. Size 1024x768.
     */
    sharePointPreviewImage?: string;
    /**
     * Defines how your tab will be made available in SharePoint.
     */
    supportedSharePointHosts?: SupportedSharePointHost[];
}

export enum ConfigurableTabContext {
    ChannelTab = "channelTab",
    MeetingChatTab = "meetingChatTab",
    MeetingDetailsTab = "meetingDetailsTab",
    MeetingSidePanel = "meetingSidePanel",
    MeetingStage = "meetingStage",
    PersonalTab = "personalTab",
    PrivateChatTab = "privateChatTab",
}

export enum MeetingSurface {
    SidePanel = "sidePanel",
    Stage = "stage",
}

export enum ConfigurableTabScope {
    GroupChat = "groupChat",
    Team = "team",
}

export enum SupportedSharePointHost {
    SharePointFullPage = "sharePointFullPage",
    SharePointWebPart = "sharePointWebPart",
}

export interface Connector {
    /**
     * The url to use for configuring the connector using the inline configuration experience.
     */
    configurationUrl?: string;
    /**
     * A unique identifier for the connector which matches its ID in the Connectors Developer
     * Portal.
     */
    connectorId: string;
    /**
     * Specifies whether the connector offers an experience in the context of a channel in a
     * team, or an experience scoped to an individual user alone. Currently, only the team scope
     * is supported.
     */
    scopes: ConnectorScope[];
}

export enum ConnectorScope {
    Team = "team",
}

export interface CopilotAgents {
    /**
     * An array of declarative agent elements references. Currently, only one declarative agent
     * per application is supported.
     */
    declarativeAgents: DeclarativeAgentRef[];
}

/**
 * A reference to a declarative agent element. The element's definition is in a separate
 * file.
 */
export interface DeclarativeAgentRef {
    /**
     * Relative file path to this declarative agent element file in the application package.
     */
    file: string;
    /**
     * A unique identifier for this declarative agent element.
     */
    id: string;
}

/**
 * Cards wich could be pinned to dashboard providing summarized view of information relevant
 * to user.
 */
export interface DashboardCard {
    contentSource: DashboardCardContentSource;
    /**
     * Rendering Size for dashboard card.
     */
    defaultSize: DefaultSize;
    /**
     * Description of the card.Maximum length is 255 characters.
     */
    description: string;
    /**
     * Represents the name of the card. Maximum length is 255 characters.
     */
    displayName: string;
    icon?:       DashboardCardIcon;
    /**
     * Unique Id for the card. Must be unique inside the app.
     */
    id: string;
    /**
     * Id of the group in the card picker. This must be guid.
     */
    pickerGroupId: string;
}

/**
 * Represents a configuration for the source of the card’s content.
 */
export interface DashboardCardContentSource {
    /**
     * The configuration for the bot source. Required if sourceType is set to bot.
     */
    botConfiguration?: BotConfiguration;
    /**
     * The content of the dashboard card is sourced from a bot.
     */
    sourceType?: SourceType;
}

/**
 * The configuration for the bot source. Required if sourceType is set to bot.
 */
export interface BotConfiguration {
    /**
     * The unique Microsoft app ID for the bot as registered with the Bot Framework.
     */
    botId?: string;
}

/**
 * The content of the dashboard card is sourced from a bot.
 */
export enum SourceType {
    Bot = "bot",
}

/**
 * Rendering Size for dashboard card.
 */
export enum DefaultSize {
    Large = "large",
    Medium = "medium",
}

/**
 * Represents a configuration for the source of the card’s content
 */
export interface DashboardCardIcon {
    /**
     * The icon for the card, to be displayed in the toolbox and card bar, represented as URL.
     */
    iconUrl?: string;
    /**
     * Office UI Fabric/Fluent UI icon friendly name for the card. This value will be used if
     * ‘iconUrl’ is not specified.
     */
    officeUIFabricIconName?: string;
}

/**
 * When a group install scope is selected, this will define the default capability when the
 * user installs the app
 */
export interface DefaultGroupCapability {
    /**
     * When the install scope selected is GroupChat, this field specifies the default capability
     * available
     */
    groupchat?: Groupchat;
    /**
     * When the install scope selected is Meetings, this field specifies the default capability
     * available
     */
    meetings?: Groupchat;
    /**
     * When the install scope selected is Team, this field specifies the default capability
     * available
     */
    team?: Groupchat;
}

/**
 * When the install scope selected is GroupChat, this field specifies the default capability
 * available
 *
 * When the install scope selected is Meetings, this field specifies the default capability
 * available
 *
 * When the install scope selected is Team, this field specifies the default capability
 * available
 */
export enum Groupchat {
    Bot = "bot",
    Connector = "connector",
    Tab = "tab",
}

/**
 * The install scope defined for this app by default. This will be the option displayed on
 * the button when a user tries to add the app
 */
export enum DefaultInstallScope {
    GroupChat = "groupChat",
    Meetings = "meetings",
    Personal = "personal",
    Team = "team",
}

export interface Description {
    /**
     * The full description of the app. Maximum length is 4000 characters.
     */
    full: string;
    /**
     * A short description of the app used when space is limited. Maximum length is 80
     * characters.
     */
    short: string;
}

export interface Developer {
    /**
     * The Microsoft Partner Network ID that identifies the partner organization building the
     * app. This field is not required, and should only be used if you are already part of the
     * Microsoft Partner Network. More info at https://aka.ms/partner
     */
    mpnId?: string;
    /**
     * The display name for the developer.
     */
    name: string;
    /**
     * The url to the page that provides privacy information for the app.
     */
    privacyUrl: string;
    /**
     * The url to the page that provides the terms of use for the app.
     */
    termsOfUseUrl: string;
    /**
     * The url to the page that provides support information for the app.
     */
    websiteUrl: string;
}

export enum DevicePermission {
    Geolocation = "geolocation",
    MIDI = "midi",
    Media = "media",
    Notifications = "notifications",
    OpenExternal = "openExternal",
}

/**
 * The set of extensions for this app. Currently only one extensions per app is supported.
 */
export interface ElementExtension {
    alternates?: ExtensionAlternateVersionsArray[];
    /**
     * The url for your extension, used to validate Exchange user identity tokens.
     */
    audienceClaimUrl?: string;
    autoRunEvents?:    ExtensionAutoRunEventsArray[];
    requirements?:     RequirementsExtensionElement;
    ribbons?:          ExtensionRibbonsArray[];
    runtimes?:         ExtensionRuntimesArray[];
}

export interface ExtensionAlternateVersionsArray {
    alternateIcons?: AlternateIcons;
    hide?:           Hide;
    prefer?:         Prefer;
    requirements?:   RequirementsExtensionElement;
}

export interface AlternateIcons {
    highResolutionIcon: ExtensionCommonIcon;
    icon:               ExtensionCommonIcon;
}

export interface ExtensionCommonIcon {
    /**
     * Size in pixels of the icon. Three image sizes are required (16, 32, and 80 pixels)
     */
    size: number;
    /**
     * Absolute Url to the icon.
     */
    url: string;
}

export interface Hide {
    customOfficeAddin?: CustomOfficeAddin;
    storeOfficeAddin?:  StoreOfficeAddin;
    [property: string]: any;
}

export interface CustomOfficeAddin {
    /**
     * Solution ID of the in-market add-in to hide. Maximum length is 64 characters.
     */
    officeAddinId: string;
}

export interface StoreOfficeAddin {
    /**
     * Asset ID of the in-market add-in to hide. Maximum length is 64 characters.
     */
    assetId: string;
    /**
     * Solution ID of an in-market add-in to hide. Maximum length is 64 characters.
     */
    officeAddinId: string;
}

export interface Prefer {
    comAddin?: COMAddin;
    [property: string]: any;
}

export interface COMAddin {
    /**
     * Program ID of the alternate com extension. Maximum length is 64 characters.
     */
    progId: string;
}

export interface RequirementsExtensionElement {
    capabilities?: Capability[];
    /**
     * Identifies the form factors that support the add-in. Supported values: mobile, desktop.
     */
    formFactors?: FormFactor[];
    /**
     * Identifies the scopes in which the add-in can run.
     */
    scopes?: RequirementsScope[];
}

export interface Capability {
    /**
     * Identifies the maximum version for the requirement sets that the add-in needs to run.
     */
    maxVersion?: string;
    /**
     * Identifies the minimum version for the requirement sets that the add-in needs to run.
     */
    minVersion?: string;
    /**
     * Identifies the name of the requirement sets that the add-in needs to run.
     */
    name: string;
}

export enum FormFactor {
    Desktop = "desktop",
    Mobile = "mobile",
}

export enum RequirementsScope {
    Document = "document",
    Mail = "mail",
    Presentation = "presentation",
    Workbook = "workbook",
}

export interface ExtensionAutoRunEventsArray {
    /**
     * Specifies the type of event. For supported types, please see:
     * https://review.learn.microsoft.com/en-us/office/dev/add-ins/outlook/autolaunch?tabs=xmlmanifest#supported-events.
     */
    events:        Event[];
    requirements?: RequirementsExtensionElement;
}

export interface Event {
    /**
     * The ID of an action defined in runtimes. Maximum length is 64 characters.
     */
    actionId: string;
    /**
     * Configures how Outlook responds to the event.
     */
    options?: Options;
    type:     string;
}

/**
 * Configures how Outlook responds to the event.
 */
export interface Options {
    sendMode: SendMode;
}

export enum SendMode {
    Block = "block",
    PromptUser = "promptUser",
    SoftBlock = "softBlock",
}

export interface ExtensionRibbonsArray {
    contexts?:     ExtensionContext[];
    requirements?: RequirementsExtensionElement;
    tabs:          ExtensionRibbonsArrayTabsItem[];
}

/**
 * Specifies the Office application windows in which the ribbon customization is available
 * to the user. Each item in the array is a member of a string array. Possible values are:
 * mailRead, mailCompose, meetingDetailsOrganizer, meetingDetailsAttendee.
 */
export enum ExtensionContext {
    Default = "default",
    LogEventMeetingDetailsAttendee = "logEventMeetingDetailsAttendee",
    MailCompose = "mailCompose",
    MailRead = "mailRead",
    MeetingDetailsAttendee = "meetingDetailsAttendee",
    MeetingDetailsOrganizer = "meetingDetailsOrganizer",
    OnlineMeetingDetailsOrganizer = "onlineMeetingDetailsOrganizer",
}

export interface ExtensionRibbonsArrayTabsItem {
    /**
     * Id of the existing office Tab. Maximum length is 64 characters.
     */
    builtInTabId?: string;
    /**
     * Defines mobile group item.
     */
    customMobileRibbonGroups?: ExtensionRibbonsCustomMobileGroupItem[];
    /**
     * Defines tab groups.
     */
    groups?: ExtensionRibbonsCustomTabGroupsItem[];
    /**
     * A unique identifier for this tab within the app. Maximum length is 64 characters.
     */
    id?: string;
    /**
     * Displayed text for the tab. Maximum length is 64 characters.
     */
    label?:    string;
    position?: Position;
}

export interface ExtensionRibbonsCustomMobileGroupItem {
    controls: ExtensionRibbonsCustomMobileControlButtonItem[];
    /**
     * Specify the Id of the group. Used for mobileMessageRead ext point.
     */
    id: string;
    /**
     * Short label of the control. Maximum length is 32 characters.
     */
    label: string;
    [property: string]: any;
}

export interface ExtensionRibbonsCustomMobileControlButtonItem {
    /**
     * The ID of an action defined in runtimes. Maximum length is 64 characters.
     */
    actionId: string;
    icons:    ExtensionCustomMobileIcon[];
    /**
     * Specify the Id of the button like msgReadFunctionButton.
     */
    id: string;
    /**
     * Short label of the control. Maximum length is 32 characters.
     */
    label: string;
    type:  PurpleType;
    [property: string]: any;
}

export interface ExtensionCustomMobileIcon {
    /**
     * How to scale - 1,2,3 for each image. This attribute specifies the UIScreen.scale property
     * for iOS devices.
     */
    scale: number;
    /**
     * Size in pixels of the icon. Three image sizes are required (25, 32, and 48 pixels).
     */
    size: number;
    /**
     * Url to the icon.
     */
    url: string;
}

export enum PurpleType {
    MobileButton = "mobileButton",
}

export interface ExtensionRibbonsCustomTabGroupsItem {
    /**
     * Id of a built-in Group. Maximum length is 64 characters.
     */
    builtInGroupId?: string;
    controls?:       ExtensionCommonCustomGroupControlsItem[];
    icons?:          ExtensionCommonIcon[];
    /**
     * A unique identifier for this group within the app. Maximum length is 64 characters.
     */
    id?: string;
    /**
     * Displayed text for the group. Maximum length is 64 characters.
     */
    label?: string;
}

export interface ExtensionCommonCustomGroupControlsItem {
    /**
     * The ID of an execution-type action that handles this key combination. Maximum length is
     * 64 characters.
     */
    actionId: string;
    /**
     * Id of the existing office control. Maximum length is 64 characters.
     */
    builtInControlId?: string;
    /**
     * Whether the control is initially enabled.
     */
    enabled?: boolean;
    icons:    ExtensionCommonIcon[];
    /**
     * A unique identifier for this control within the app. Maximum length is 64 characters.
     */
    id: string;
    /**
     * Configures the items for a menu control.
     */
    items?: ExtensionCommonCustomControlMenuItem[];
    /**
     * Displayed text for the control. Maximum length is 64 characters.
     */
    label: string;
    /**
     * Specifies whether a group, button, menu, or menu item will be hidden on application and
     * platform combinations that support the API (Office.ribbon.requestCreateControls) that
     * installs custom contextual tabs on the ribbon. Default is false.
     */
    overriddenByRibbonApi?: boolean;
    supertip:               ExtensionCommonSuperToolTip;
    /**
     * Defines the type of control whether button or menu.
     */
    type: FluffyType;
}

export interface ExtensionCommonCustomControlMenuItem {
    /**
     * The ID of an action defined in runtimes. Maximum length is 64 characters.
     */
    actionId: string;
    /**
     * Whether the control is initially enabled.
     */
    enabled?: boolean;
    icons?:   ExtensionCommonIcon[];
    /**
     * A unique identifier for this control within the app. Maximum length is 64 characters.
     */
    id: string;
    /**
     * Displayed text for the control. Maximum length is 64 characters.
     */
    label:                  string;
    overriddenByRibbonApi?: boolean;
    supertip:               ExtensionCommonSuperToolTip;
    /**
     * Supported values: menuItem.
     */
    type: ItemType;
}

export interface ExtensionCommonSuperToolTip {
    /**
     * Description of the super tip. Maximum length is 250 characters.
     */
    description: string;
    /**
     * Title text of the super tip. Maximum length is 64 characters.
     */
    title: string;
}

/**
 * Supported values: menuItem.
 */
export enum ItemType {
    MenuItem = "menuItem",
}

/**
 * Defines the type of control whether button or menu.
 */
export enum FluffyType {
    Button = "button",
    Menu = "menu",
}

export interface Position {
    /**
     * Define alignment of this custom tab relative to the specified built-in tab.
     */
    align: Align;
    /**
     * The id of the built-in tab. Maximum length is 64 characters.
     */
    builtInTabId: string;
}

/**
 * Define alignment of this custom tab relative to the specified built-in tab.
 */
export enum Align {
    After = "after",
    Before = "before",
}

/**
 * A runtime environment for a page or script
 */
export interface ExtensionRuntimesArray {
    actions?: ExtensionRuntimesActionsItem[];
    code:     ExtensionRuntimeCode;
    /**
     * A unique identifier for this runtime within the app.  Maximum length is 64 characters.
     */
    id: string;
    /**
     * Runtimes with a short lifetime do not preserve state across executions. Runtimes with a
     * long lifetime do.
     */
    lifetime?:     Lifetime;
    requirements?: RequirementsExtensionElement;
    /**
     * Supports running functions and launching pages.
     */
    type?: RuntimeType;
}

/**
 * Specifies the set of actions supported by this runtime. An action is either running a
 * JavaScript function or opening a view such as a task pane.
 */
export interface ExtensionRuntimesActionsItem {
    /**
     * Display name of the action. Maximum length is 64 characters.
     */
    displayName?: string;
    /**
     * Identifier for this action. Maximum length is 64 characters. This value is passed to the
     * code file.
     */
    id: string;
    /**
     * Whether allows the action to have multiple selection.
     */
    multiselect?: boolean;
    /**
     * Specifies that a task pane supports pinning, which keeps the task pane open when the user
     * changes the selection.
     */
    pinnable?: boolean;
    /**
     * Whether allows task pane add-ins to activate without the Reading Pane enabled or a
     * message selected.
     */
    supportsNoItemContext?: boolean;
    /**
     * executeFunction: Run a script function without waiting for it to finish. openPate: Open a
     * page in a view.
     */
    type: ActionType;
    /**
     * View where the page should be opened. Maximum length is 64 characters.
     */
    view?: string;
}

/**
 * executeFunction: Run a script function without waiting for it to finish. openPate: Open a
 * page in a view.
 */
export enum ActionType {
    ExecuteFunction = "executeFunction",
    OpenPage = "openPage",
}

export interface ExtensionRuntimeCode {
    /**
     * URL of the .html page to be loaded in browser-based runtimes.
     */
    page: string;
    /**
     * URL of the .js script file to be loaded in UI-less runtimes.
     */
    script?: string;
}

/**
 * Runtimes with a short lifetime do not preserve state across executions. Runtimes with a
 * long lifetime do.
 */
export enum Lifetime {
    Long = "long",
    Short = "short",
}

/**
 * Supports running functions and launching pages.
 */
export enum RuntimeType {
    General = "general",
}

/**
 * Specify the app's Graph connector configuration. If this is present then
 * webApplicationInfo.id must also be specified.
 */
export interface GraphConnector {
    /**
     * The url where Graph-connector notifications for the application should be sent.
     */
    notificationUrl: string;
}

export interface Icons {
    /**
     * A relative file path to a full color PNG icon. Size 192x192.
     */
    color: string;
    /**
     * A relative file path to a transparent PNG outline icon. The border color needs to be
     * white. Size 32x32.
     */
    outline: string;
}

export interface LocalizationInfo {
    additionalLanguages?: AdditionalLanguage[];
    /**
     * A relative file path to a the .json file containing strings in the default language.
     */
    defaultLanguageFile?: string;
    /**
     * The language tag of the strings in this top level manifest file.
     */
    defaultLanguageTag: string;
}

export interface AdditionalLanguage {
    /**
     * A relative file path to a the .json file containing the translated strings.
     */
    file: string;
    /**
     * The language tag of the strings in the provided file.
     */
    languageTag: string;
}

export enum ManifestVersion {
    The119 = "1.19",
}

/**
 * Specify meeting extension definition.
 */
export interface MeetingExtensionDefinition {
    /**
     * Meeting supported scenes.
     */
    scenes?: Scene[];
    /**
     * A boolean value indicating whether this app allows management by anonymous users.
     */
    supportsAnonymousGuestUsers?: boolean;
    /**
     * A boolean value indicating whether this app can stream the meeting's audio video content
     * to an RTMP endpoint.
     */
    supportsStreaming?: boolean;
}

export interface Scene {
    /**
     * A relative file path to a scene metadata json file.
     */
    file: string;
    /**
     * A unique identifier for this scene. This id must be a GUID.
     */
    id: string;
    /**
     * Maximum audiences supported in scene.
     */
    maxAudience: number;
    /**
     * Scene name.
     */
    name: string;
    /**
     * A relative file path to a scene PNG preview icon.
     */
    preview: string;
    /**
     * Number of seats reserved for organizers or presenters.
     */
    seatsReservedForOrganizersOrPresenters: number;
}

export interface Name {
    /**
     * The full name of the app, used if the full app name exceeds 30 characters.
     */
    full: string;
    /**
     * A short display name for the app.
     */
    short: string;
}

export enum Permission {
    Identity = "identity",
    MessageTeamMembers = "messageTeamMembers",
}

export interface StaticTab {
    /**
     * The Microsoft App ID specified for the bot in the Bot Framework portal
     * (https://dev.botframework.com/bots)
     */
    contentBotId?: string;
    /**
     * The url which points to the entity UI to be displayed in the canvas.
     */
    contentUrl?: string;
    /**
     * The set of contextItem scopes that a tab belong to
     */
    context?: StaticTabContext[];
    /**
     * A unique identifier for the entity which the tab displays.
     */
    entityId: string;
    /**
     * The display name of the tab.
     */
    name?: string;
    /**
     * Specifies whether the tab offers an experience in the context of a channel in a team, or
     * an experience scoped to an individual user alone or group chat. These options are
     * non-exclusive. Currently static tabs are only supported in the 'personal' scope.
     */
    scopes: CommandListScope[];
    /**
     * The url to direct a user's search queries.
     */
    searchUrl?: string;
    /**
     * The url to point at if a user opts to view in a browser.
     */
    websiteUrl?: string;
}

export enum StaticTabContext {
    ChannelTab = "channelTab",
    MeetingChatTab = "meetingChatTab",
    MeetingDetailsTab = "meetingDetailsTab",
    MeetingSidePanel = "meetingSidePanel",
    MeetingStage = "meetingStage",
    PersonalTab = "personalTab",
    PrivateChatTab = "privateChatTab",
    TeamLevelApp = "teamLevelApp",
}

/**
 * Subscription offer associated with this app.
 */
export interface SubscriptionOffer {
    /**
     * A unique identifier for the Commercial Marketplace Software as a Service Offer.
     */
    offerId: string;
}

export enum SupportedChannelType {
    PrivateChannels = "privateChannels",
    SharedChannels = "sharedChannels",
}

/**
 * Specify your AAD App ID and Graph information to help users seamlessly sign into your AAD
 * app.
 */
export interface WebApplicationInfo {
    /**
     * AAD application id of the app. This id must be a GUID.
     */
    id: string;
    /**
     * Resource url of app for acquiring auth token for SSO.
     */
    resource?: string;
}

// Converts JSON strings to/from your types
// and asserts the results of JSON.parse at runtime
export class Convert {
    public static toMicrosoftTeamsV1D19(json: string): MicrosoftTeamsV1D19 {
        return cast(JSON.parse(json), r("MicrosoftTeamsV1D19"));
    }

    public static microsoftTeamsV1D19ToJson(value: MicrosoftTeamsV1D19): string {
        return JSON.stringify(uncast(value, r("MicrosoftTeamsV1D19")), null, 2);
    }
}

function invalidValue(typ: any, val: any, key: any, parent: any = ''): never {
    const prettyTyp = prettyTypeName(typ);
    const parentText = parent ? ` on ${parent}` : '';
    const keyText = key ? ` for key "${key}"` : '';
    throw Error(`Invalid value${keyText}${parentText}. Expected ${prettyTyp} but got ${JSON.stringify(val)}`);
}

function prettyTypeName(typ: any): string {
    if (Array.isArray(typ)) {
        if (typ.length === 2 && typ[0] === undefined) {
            return `an optional ${prettyTypeName(typ[1])}`;
        } else {
            return `one of [${typ.map(a => { return prettyTypeName(a); }).join(", ")}]`;
        }
    } else if (typeof typ === "object" && typ.literal !== undefined) {
        return typ.literal;
    } else {
        return typeof typ;
    }
}

function jsonToJSProps(typ: any): any {
    if (typ.jsonToJS === undefined) {
        const map: any = {};
        typ.props.forEach((p: any) => map[p.json] = { key: p.js, typ: p.typ });
        typ.jsonToJS = map;
    }
    return typ.jsonToJS;
}

function jsToJSONProps(typ: any): any {
    if (typ.jsToJSON === undefined) {
        const map: any = {};
        typ.props.forEach((p: any) => map[p.js] = { key: p.json, typ: p.typ });
        typ.jsToJSON = map;
    }
    return typ.jsToJSON;
}

function transform(val: any, typ: any, getProps: any, key: any = '', parent: any = ''): any {
    function transformPrimitive(typ: string, val: any): any {
        if (typeof typ === typeof val) return val;
        return invalidValue(typ, val, key, parent);
    }

    function transformUnion(typs: any[], val: any): any {
        // val must validate against one typ in typs
        const l = typs.length;
        for (let i = 0; i < l; i++) {
            const typ = typs[i];
            try {
                return transform(val, typ, getProps);
            } catch (_) {}
        }
        return invalidValue(typs, val, key, parent);
    }

    function transformEnum(cases: string[], val: any): any {
        if (cases.indexOf(val) !== -1) return val;
        return invalidValue(cases.map(a => { return l(a); }), val, key, parent);
    }

    function transformArray(typ: any, val: any): any {
        // val must be an array with no invalid elements
        if (!Array.isArray(val)) return invalidValue(l("array"), val, key, parent);
        return val.map(el => transform(el, typ, getProps));
    }

    function transformDate(val: any): any {
        if (val === null) {
            return null;
        }
        const d = new Date(val);
        if (isNaN(d.valueOf())) {
            return invalidValue(l("Date"), val, key, parent);
        }
        return d;
    }

    function transformObject(props: { [k: string]: any }, additional: any, val: any): any {
        if (val === null || typeof val !== "object" || Array.isArray(val)) {
            return invalidValue(l(ref || "object"), val, key, parent);
        }
        const result: any = {};
        Object.getOwnPropertyNames(props).forEach(key => {
            const prop = props[key];
            const v = Object.prototype.hasOwnProperty.call(val, key) ? val[key] : undefined;
            result[prop.key] = transform(v, prop.typ, getProps, key, ref);
        });
        Object.getOwnPropertyNames(val).forEach(key => {
            if (!Object.prototype.hasOwnProperty.call(props, key)) {
                result[key] = transform(val[key], additional, getProps, key, ref);
            }
        });
        return result;
    }

    if (typ === "any") return val;
    if (typ === null) {
        if (val === null) return val;
        return invalidValue(typ, val, key, parent);
    }
    if (typ === false) return invalidValue(typ, val, key, parent);
    let ref: any = undefined;
    while (typeof typ === "object" && typ.ref !== undefined) {
        ref = typ.ref;
        typ = typeMap[typ.ref];
    }
    if (Array.isArray(typ)) return transformEnum(typ, val);
    if (typeof typ === "object") {
        return typ.hasOwnProperty("unionMembers") ? transformUnion(typ.unionMembers, val)
            : typ.hasOwnProperty("arrayItems")    ? transformArray(typ.arrayItems, val)
            : typ.hasOwnProperty("props")         ? transformObject(getProps(typ), typ.additional, val)
            : invalidValue(typ, val, key, parent);
    }
    // Numbers can be parsed by Date but shouldn't be.
    if (typ === Date && typeof val !== "number") return transformDate(val);
    return transformPrimitive(typ, val);
}

function cast<T>(val: any, typ: any): T {
    return transform(val, typ, jsonToJSProps);
}

function uncast<T>(val: T, typ: any): any {
    return transform(val, typ, jsToJSONProps);
}

function l(typ: any) {
    return { literal: typ };
}

function a(typ: any) {
    return { arrayItems: typ };
}

function u(...typs: any[]) {
    return { unionMembers: typs };
}

function o(props: any[], additional: any) {
    return { props, additional };
}

function m(additional: any) {
    return { props: [], additional };
}

function r(name: string) {
    return { ref: name };
}

const typeMap: any = {
    "MicrosoftTeamsV1D19": o([
        { json: "$schema", js: "$schema", typ: u(undefined, "") },
        { json: "accentColor", js: "accentColor", typ: "" },
        { json: "activities", js: "activities", typ: u(undefined, r("Activities")) },
        { json: "authorization", js: "authorization", typ: u(undefined, r("MicrosoftTeamsV1D19Authorization")) },
        { json: "bots", js: "bots", typ: u(undefined, a(r("Bot"))) },
        { json: "composeExtensions", js: "composeExtensions", typ: u(undefined, a(r("ComposeExtension"))) },
        { json: "configurableProperties", js: "configurableProperties", typ: u(undefined, a(r("ConfigurableProperty"))) },
        { json: "configurableTabs", js: "configurableTabs", typ: u(undefined, a(r("ConfigurableTab"))) },
        { json: "connectors", js: "connectors", typ: u(undefined, a(r("Connector"))) },
        { json: "copilotAgents", js: "copilotAgents", typ: u(undefined, r("CopilotAgents")) },
        { json: "dashboardCards", js: "dashboardCards", typ: u(undefined, a(r("DashboardCard"))) },
        { json: "defaultBlockUntilAdminAction", js: "defaultBlockUntilAdminAction", typ: u(undefined, true) },
        { json: "defaultGroupCapability", js: "defaultGroupCapability", typ: u(undefined, r("DefaultGroupCapability")) },
        { json: "defaultInstallScope", js: "defaultInstallScope", typ: u(undefined, r("DefaultInstallScope")) },
        { json: "description", js: "description", typ: r("Description") },
        { json: "developer", js: "developer", typ: r("Developer") },
        { json: "devicePermissions", js: "devicePermissions", typ: u(undefined, a(r("DevicePermission"))) },
        { json: "extensions", js: "extensions", typ: u(undefined, a(r("ElementExtension"))) },
        { json: "graphConnector", js: "graphConnector", typ: u(undefined, r("GraphConnector")) },
        { json: "icons", js: "icons", typ: r("Icons") },
        { json: "id", js: "id", typ: "" },
        { json: "isFullScreen", js: "isFullScreen", typ: u(undefined, true) },
        { json: "localizationInfo", js: "localizationInfo", typ: u(undefined, r("LocalizationInfo")) },
        { json: "manifestVersion", js: "manifestVersion", typ: r("ManifestVersion") },
        { json: "meetingExtensionDefinition", js: "meetingExtensionDefinition", typ: u(undefined, r("MeetingExtensionDefinition")) },
        { json: "name", js: "name", typ: r("Name") },
        { json: "permissions", js: "permissions", typ: u(undefined, a(r("Permission"))) },
        { json: "publisherDocsUrl", js: "publisherDocsUrl", typ: u(undefined, "") },
        { json: "showLoadingIndicator", js: "showLoadingIndicator", typ: u(undefined, true) },
        { json: "staticTabs", js: "staticTabs", typ: u(undefined, a(r("StaticTab"))) },
        { json: "subscriptionOffer", js: "subscriptionOffer", typ: u(undefined, r("SubscriptionOffer")) },
        { json: "supportedChannelTypes", js: "supportedChannelTypes", typ: u(undefined, a(r("SupportedChannelType"))) },
        { json: "validDomains", js: "validDomains", typ: u(undefined, a("")) },
        { json: "version", js: "version", typ: "" },
        { json: "webApplicationInfo", js: "webApplicationInfo", typ: u(undefined, r("WebApplicationInfo")) },
    ], false),
    "Activities": o([
        { json: "activityTypes", js: "activityTypes", typ: u(undefined, a(r("ActivityType"))) },
    ], false),
    "ActivityType": o([
        { json: "description", js: "description", typ: "" },
        { json: "templateText", js: "templateText", typ: "" },
        { json: "type", js: "type", typ: "" },
    ], false),
    "MicrosoftTeamsV1D19Authorization": o([
        { json: "permissions", js: "permissions", typ: u(undefined, r("Permissions")) },
    ], false),
    "Permissions": o([
        { json: "resourceSpecific", js: "resourceSpecific", typ: u(undefined, a(r("ResourceSpecific"))) },
    ], false),
    "ResourceSpecific": o([
        { json: "name", js: "name", typ: "" },
        { json: "type", js: "type", typ: r("ResourceSpecificType") },
    ], false),
    "Bot": o([
        { json: "botId", js: "botId", typ: "" },
        { json: "commandLists", js: "commandLists", typ: u(undefined, a(r("CommandList"))) },
        { json: "configuration", js: "configuration", typ: u(undefined, r("Configuration")) },
        { json: "isNotificationOnly", js: "isNotificationOnly", typ: u(undefined, true) },
        { json: "needsChannelSelector", js: "needsChannelSelector", typ: u(undefined, true) },
        { json: "scopes", js: "scopes", typ: a(r("CommandListScope")) },
        { json: "supportsCalling", js: "supportsCalling", typ: u(undefined, true) },
        { json: "supportsFiles", js: "supportsFiles", typ: u(undefined, true) },
        { json: "supportsVideo", js: "supportsVideo", typ: u(undefined, true) },
    ], false),
    "CommandList": o([
        { json: "commands", js: "commands", typ: a(r("CommandListCommand")) },
        { json: "scopes", js: "scopes", typ: a(r("CommandListScope")) },
    ], false),
    "CommandListCommand": o([
        { json: "description", js: "description", typ: "" },
        { json: "title", js: "title", typ: "" },
    ], false),
    "Configuration": o([
        { json: "groupChat", js: "groupChat", typ: u(undefined, r("GroupChat")) },
        { json: "team", js: "team", typ: u(undefined, r("GroupChat")) },
    ], false),
    "GroupChat": o([
        { json: "fetchTask", js: "fetchTask", typ: u(undefined, true) },
        { json: "taskInfo", js: "taskInfo", typ: u(undefined, r("TaskInfo")) },
    ], false),
    "TaskInfo": o([
        { json: "height", js: "height", typ: u(undefined, "") },
        { json: "title", js: "title", typ: u(undefined, "") },
        { json: "url", js: "url", typ: u(undefined, "") },
        { json: "width", js: "width", typ: u(undefined, "") },
    ], false),
    "ComposeExtension": o([
        { json: "apiSpecificationFile", js: "apiSpecificationFile", typ: u(undefined, "") },
        { json: "authorization", js: "authorization", typ: u(undefined, r("ComposeExtensionAuthorization")) },
        { json: "botId", js: "botId", typ: u(undefined, "") },
        { json: "canUpdateConfiguration", js: "canUpdateConfiguration", typ: u(undefined, u(true, null)) },
        { json: "commands", js: "commands", typ: u(undefined, a(r("ComposeExtensionCommand"))) },
        { json: "composeExtensionType", js: "composeExtensionType", typ: u(undefined, r("ComposeExtensionType")) },
        { json: "messageHandlers", js: "messageHandlers", typ: u(undefined, a(r("MessageHandler"))) },
    ], false),
    "ComposeExtensionAuthorization": o([
        { json: "apiSecretServiceAuthConfiguration", js: "apiSecretServiceAuthConfiguration", typ: u(undefined, r("APISecretServiceAuthConfiguration")) },
        { json: "authType", js: "authType", typ: u(undefined, r("AuthType")) },
        { json: "microsoftEntraConfiguration", js: "microsoftEntraConfiguration", typ: u(undefined, r("MicrosoftEntraConfiguration")) },
    ], false),
    "APISecretServiceAuthConfiguration": o([
        { json: "apiSecretRegistrationId", js: "apiSecretRegistrationId", typ: u(undefined, "") },
    ], false),
    "MicrosoftEntraConfiguration": o([
        { json: "supportsSingleSignOn", js: "supportsSingleSignOn", typ: u(undefined, true) },
    ], false),
    "ComposeExtensionCommand": o([
        { json: "apiResponseRenderingTemplateFile", js: "apiResponseRenderingTemplateFile", typ: u(undefined, "") },
        { json: "context", js: "context", typ: u(undefined, a(r("CommandContext"))) },
        { json: "description", js: "description", typ: u(undefined, "") },
        { json: "fetchTask", js: "fetchTask", typ: u(undefined, true) },
        { json: "id", js: "id", typ: "" },
        { json: "initialRun", js: "initialRun", typ: u(undefined, true) },
        { json: "parameters", js: "parameters", typ: u(undefined, a(r("Parameter"))) },
        { json: "samplePrompts", js: "samplePrompts", typ: u(undefined, a(r("SamplePrompt"))) },
        { json: "semanticDescription", js: "semanticDescription", typ: u(undefined, "") },
        { json: "taskInfo", js: "taskInfo", typ: u(undefined, r("TaskInfo")) },
        { json: "title", js: "title", typ: "" },
        { json: "type", js: "type", typ: u(undefined, r("CommandType")) },
    ], false),
    "Parameter": o([
        { json: "choices", js: "choices", typ: u(undefined, a(r("Choice"))) },
        { json: "description", js: "description", typ: u(undefined, "") },
        { json: "inputType", js: "inputType", typ: u(undefined, r("InputType")) },
        { json: "isRequired", js: "isRequired", typ: u(undefined, true) },
        { json: "name", js: "name", typ: "" },
        { json: "semanticDescription", js: "semanticDescription", typ: u(undefined, "") },
        { json: "title", js: "title", typ: "" },
        { json: "value", js: "value", typ: u(undefined, "") },
    ], false),
    "Choice": o([
        { json: "title", js: "title", typ: "" },
        { json: "value", js: "value", typ: "" },
    ], false),
    "SamplePrompt": o([
        { json: "text", js: "text", typ: "" },
    ], false),
    "MessageHandler": o([
        { json: "type", js: "type", typ: r("MessageHandlerType") },
        { json: "value", js: "value", typ: r("Value") },
    ], false),
    "Value": o([
        { json: "domains", js: "domains", typ: u(undefined, a("")) },
        { json: "supportsAnonymizedPayloads", js: "supportsAnonymizedPayloads", typ: u(undefined, true) },
    ], "any"),
    "ConfigurableTab": o([
        { json: "canUpdateConfiguration", js: "canUpdateConfiguration", typ: u(undefined, true) },
        { json: "configurationUrl", js: "configurationUrl", typ: "" },
        { json: "context", js: "context", typ: u(undefined, a(r("ConfigurableTabContext"))) },
        { json: "meetingSurfaces", js: "meetingSurfaces", typ: u(undefined, a(r("MeetingSurface"))) },
        { json: "scopes", js: "scopes", typ: a(r("ConfigurableTabScope")) },
        { json: "sharePointPreviewImage", js: "sharePointPreviewImage", typ: u(undefined, "") },
        { json: "supportedSharePointHosts", js: "supportedSharePointHosts", typ: u(undefined, a(r("SupportedSharePointHost"))) },
    ], false),
    "Connector": o([
        { json: "configurationUrl", js: "configurationUrl", typ: u(undefined, "") },
        { json: "connectorId", js: "connectorId", typ: "" },
        { json: "scopes", js: "scopes", typ: a(r("ConnectorScope")) },
    ], false),
    "CopilotAgents": o([
        { json: "declarativeAgents", js: "declarativeAgents", typ: a(r("DeclarativeAgentRef")) },
    ], false),
    "DeclarativeAgentRef": o([
        { json: "file", js: "file", typ: "" },
        { json: "id", js: "id", typ: "" },
    ], false),
    "DashboardCard": o([
        { json: "contentSource", js: "contentSource", typ: r("DashboardCardContentSource") },
        { json: "defaultSize", js: "defaultSize", typ: r("DefaultSize") },
        { json: "description", js: "description", typ: "" },
        { json: "displayName", js: "displayName", typ: "" },
        { json: "icon", js: "icon", typ: u(undefined, r("DashboardCardIcon")) },
        { json: "id", js: "id", typ: "" },
        { json: "pickerGroupId", js: "pickerGroupId", typ: "" },
    ], false),
    "DashboardCardContentSource": o([
        { json: "botConfiguration", js: "botConfiguration", typ: u(undefined, r("BotConfiguration")) },
        { json: "sourceType", js: "sourceType", typ: u(undefined, r("SourceType")) },
    ], false),
    "BotConfiguration": o([
        { json: "botId", js: "botId", typ: u(undefined, "") },
    ], false),
    "DashboardCardIcon": o([
        { json: "iconUrl", js: "iconUrl", typ: u(undefined, "") },
        { json: "officeUIFabricIconName", js: "officeUIFabricIconName", typ: u(undefined, "") },
    ], false),
    "DefaultGroupCapability": o([
        { json: "groupchat", js: "groupchat", typ: u(undefined, r("Groupchat")) },
        { json: "meetings", js: "meetings", typ: u(undefined, r("Groupchat")) },
        { json: "team", js: "team", typ: u(undefined, r("Groupchat")) },
    ], false),
    "Description": o([
        { json: "full", js: "full", typ: "" },
        { json: "short", js: "short", typ: "" },
    ], false),
    "Developer": o([
        { json: "mpnId", js: "mpnId", typ: u(undefined, "") },
        { json: "name", js: "name", typ: "" },
        { json: "privacyUrl", js: "privacyUrl", typ: "" },
        { json: "termsOfUseUrl", js: "termsOfUseUrl", typ: "" },
        { json: "websiteUrl", js: "websiteUrl", typ: "" },
    ], false),
    "ElementExtension": o([
        { json: "alternates", js: "alternates", typ: u(undefined, a(r("ExtensionAlternateVersionsArray"))) },
        { json: "audienceClaimUrl", js: "audienceClaimUrl", typ: u(undefined, "") },
        { json: "autoRunEvents", js: "autoRunEvents", typ: u(undefined, a(r("ExtensionAutoRunEventsArray"))) },
        { json: "requirements", js: "requirements", typ: u(undefined, r("RequirementsExtensionElement")) },
        { json: "ribbons", js: "ribbons", typ: u(undefined, a(r("ExtensionRibbonsArray"))) },
        { json: "runtimes", js: "runtimes", typ: u(undefined, a(r("ExtensionRuntimesArray"))) },
    ], false),
    "ExtensionAlternateVersionsArray": o([
        { json: "alternateIcons", js: "alternateIcons", typ: u(undefined, r("AlternateIcons")) },
        { json: "hide", js: "hide", typ: u(undefined, r("Hide")) },
        { json: "prefer", js: "prefer", typ: u(undefined, r("Prefer")) },
        { json: "requirements", js: "requirements", typ: u(undefined, r("RequirementsExtensionElement")) },
    ], false),
    "AlternateIcons": o([
        { json: "highResolutionIcon", js: "highResolutionIcon", typ: r("ExtensionCommonIcon") },
        { json: "icon", js: "icon", typ: r("ExtensionCommonIcon") },
    ], false),
    "ExtensionCommonIcon": o([
        { json: "size", js: "size", typ: 3.14 },
        { json: "url", js: "url", typ: "" },
    ], false),
    "Hide": o([
        { json: "customOfficeAddin", js: "customOfficeAddin", typ: u(undefined, r("CustomOfficeAddin")) },
        { json: "storeOfficeAddin", js: "storeOfficeAddin", typ: u(undefined, r("StoreOfficeAddin")) },
    ], "any"),
    "CustomOfficeAddin": o([
        { json: "officeAddinId", js: "officeAddinId", typ: "" },
    ], false),
    "StoreOfficeAddin": o([
        { json: "assetId", js: "assetId", typ: "" },
        { json: "officeAddinId", js: "officeAddinId", typ: "" },
    ], false),
    "Prefer": o([
        { json: "comAddin", js: "comAddin", typ: u(undefined, r("COMAddin")) },
    ], "any"),
    "COMAddin": o([
        { json: "progId", js: "progId", typ: "" },
    ], false),
    "RequirementsExtensionElement": o([
        { json: "capabilities", js: "capabilities", typ: u(undefined, a(r("Capability"))) },
        { json: "formFactors", js: "formFactors", typ: u(undefined, a(r("FormFactor"))) },
        { json: "scopes", js: "scopes", typ: u(undefined, a(r("RequirementsScope"))) },
    ], false),
    "Capability": o([
        { json: "maxVersion", js: "maxVersion", typ: u(undefined, "") },
        { json: "minVersion", js: "minVersion", typ: u(undefined, "") },
        { json: "name", js: "name", typ: "" },
    ], false),
    "ExtensionAutoRunEventsArray": o([
        { json: "events", js: "events", typ: a(r("Event")) },
        { json: "requirements", js: "requirements", typ: u(undefined, r("RequirementsExtensionElement")) },
    ], false),
    "Event": o([
        { json: "actionId", js: "actionId", typ: "" },
        { json: "options", js: "options", typ: u(undefined, r("Options")) },
        { json: "type", js: "type", typ: "" },
    ], false),
    "Options": o([
        { json: "sendMode", js: "sendMode", typ: r("SendMode") },
    ], false),
    "ExtensionRibbonsArray": o([
        { json: "contexts", js: "contexts", typ: u(undefined, a(r("ExtensionContext"))) },
        { json: "requirements", js: "requirements", typ: u(undefined, r("RequirementsExtensionElement")) },
        { json: "tabs", js: "tabs", typ: a(r("ExtensionRibbonsArrayTabsItem")) },
    ], false),
    "ExtensionRibbonsArrayTabsItem": o([
        { json: "builtInTabId", js: "builtInTabId", typ: u(undefined, "") },
        { json: "customMobileRibbonGroups", js: "customMobileRibbonGroups", typ: u(undefined, a(r("ExtensionRibbonsCustomMobileGroupItem"))) },
        { json: "groups", js: "groups", typ: u(undefined, a(r("ExtensionRibbonsCustomTabGroupsItem"))) },
        { json: "id", js: "id", typ: u(undefined, "") },
        { json: "label", js: "label", typ: u(undefined, "") },
        { json: "position", js: "position", typ: u(undefined, r("Position")) },
    ], false),
    "ExtensionRibbonsCustomMobileGroupItem": o([
        { json: "controls", js: "controls", typ: a(r("ExtensionRibbonsCustomMobileControlButtonItem")) },
        { json: "id", js: "id", typ: "" },
        { json: "label", js: "label", typ: "" },
    ], "any"),
    "ExtensionRibbonsCustomMobileControlButtonItem": o([
        { json: "actionId", js: "actionId", typ: "" },
        { json: "icons", js: "icons", typ: a(r("ExtensionCustomMobileIcon")) },
        { json: "id", js: "id", typ: "" },
        { json: "label", js: "label", typ: "" },
        { json: "type", js: "type", typ: r("PurpleType") },
    ], "any"),
    "ExtensionCustomMobileIcon": o([
        { json: "scale", js: "scale", typ: 3.14 },
        { json: "size", js: "size", typ: 3.14 },
        { json: "url", js: "url", typ: "" },
    ], false),
    "ExtensionRibbonsCustomTabGroupsItem": o([
        { json: "builtInGroupId", js: "builtInGroupId", typ: u(undefined, "") },
        { json: "controls", js: "controls", typ: u(undefined, a(r("ExtensionCommonCustomGroupControlsItem"))) },
        { json: "icons", js: "icons", typ: u(undefined, a(r("ExtensionCommonIcon"))) },
        { json: "id", js: "id", typ: u(undefined, "") },
        { json: "label", js: "label", typ: u(undefined, "") },
    ], false),
    "ExtensionCommonCustomGroupControlsItem": o([
        { json: "actionId", js: "actionId", typ: "" },
        { json: "builtInControlId", js: "builtInControlId", typ: u(undefined, "") },
        { json: "enabled", js: "enabled", typ: u(undefined, true) },
        { json: "icons", js: "icons", typ: a(r("ExtensionCommonIcon")) },
        { json: "id", js: "id", typ: "" },
        { json: "items", js: "items", typ: u(undefined, a(r("ExtensionCommonCustomControlMenuItem"))) },
        { json: "label", js: "label", typ: "" },
        { json: "overriddenByRibbonApi", js: "overriddenByRibbonApi", typ: u(undefined, true) },
        { json: "supertip", js: "supertip", typ: r("ExtensionCommonSuperToolTip") },
        { json: "type", js: "type", typ: r("FluffyType") },
    ], false),
    "ExtensionCommonCustomControlMenuItem": o([
        { json: "actionId", js: "actionId", typ: "" },
        { json: "enabled", js: "enabled", typ: u(undefined, true) },
        { json: "icons", js: "icons", typ: u(undefined, a(r("ExtensionCommonIcon"))) },
        { json: "id", js: "id", typ: "" },
        { json: "label", js: "label", typ: "" },
        { json: "overriddenByRibbonApi", js: "overriddenByRibbonApi", typ: u(undefined, true) },
        { json: "supertip", js: "supertip", typ: r("ExtensionCommonSuperToolTip") },
        { json: "type", js: "type", typ: r("ItemType") },
    ], false),
    "ExtensionCommonSuperToolTip": o([
        { json: "description", js: "description", typ: "" },
        { json: "title", js: "title", typ: "" },
    ], false),
    "Position": o([
        { json: "align", js: "align", typ: r("Align") },
        { json: "builtInTabId", js: "builtInTabId", typ: "" },
    ], false),
    "ExtensionRuntimesArray": o([
        { json: "actions", js: "actions", typ: u(undefined, a(r("ExtensionRuntimesActionsItem"))) },
        { json: "code", js: "code", typ: r("ExtensionRuntimeCode") },
        { json: "id", js: "id", typ: "" },
        { json: "lifetime", js: "lifetime", typ: u(undefined, r("Lifetime")) },
        { json: "requirements", js: "requirements", typ: u(undefined, r("RequirementsExtensionElement")) },
        { json: "type", js: "type", typ: u(undefined, r("RuntimeType")) },
    ], false),
    "ExtensionRuntimesActionsItem": o([
        { json: "displayName", js: "displayName", typ: u(undefined, "") },
        { json: "id", js: "id", typ: "" },
        { json: "multiselect", js: "multiselect", typ: u(undefined, true) },
        { json: "pinnable", js: "pinnable", typ: u(undefined, true) },
        { json: "supportsNoItemContext", js: "supportsNoItemContext", typ: u(undefined, true) },
        { json: "type", js: "type", typ: r("ActionType") },
        { json: "view", js: "view", typ: u(undefined, "") },
    ], false),
    "ExtensionRuntimeCode": o([
        { json: "page", js: "page", typ: "" },
        { json: "script", js: "script", typ: u(undefined, "") },
    ], false),
    "GraphConnector": o([
        { json: "notificationUrl", js: "notificationUrl", typ: "" },
    ], false),
    "Icons": o([
        { json: "color", js: "color", typ: "" },
        { json: "outline", js: "outline", typ: "" },
    ], false),
    "LocalizationInfo": o([
        { json: "additionalLanguages", js: "additionalLanguages", typ: u(undefined, a(r("AdditionalLanguage"))) },
        { json: "defaultLanguageFile", js: "defaultLanguageFile", typ: u(undefined, "") },
        { json: "defaultLanguageTag", js: "defaultLanguageTag", typ: "" },
    ], false),
    "AdditionalLanguage": o([
        { json: "file", js: "file", typ: "" },
        { json: "languageTag", js: "languageTag", typ: "" },
    ], false),
    "MeetingExtensionDefinition": o([
        { json: "scenes", js: "scenes", typ: u(undefined, a(r("Scene"))) },
        { json: "supportsAnonymousGuestUsers", js: "supportsAnonymousGuestUsers", typ: u(undefined, true) },
        { json: "supportsStreaming", js: "supportsStreaming", typ: u(undefined, true) },
    ], false),
    "Scene": o([
        { json: "file", js: "file", typ: "" },
        { json: "id", js: "id", typ: "" },
        { json: "maxAudience", js: "maxAudience", typ: 0 },
        { json: "name", js: "name", typ: "" },
        { json: "preview", js: "preview", typ: "" },
        { json: "seatsReservedForOrganizersOrPresenters", js: "seatsReservedForOrganizersOrPresenters", typ: 0 },
    ], false),
    "Name": o([
        { json: "full", js: "full", typ: "" },
        { json: "short", js: "short", typ: "" },
    ], false),
    "StaticTab": o([
        { json: "contentBotId", js: "contentBotId", typ: u(undefined, "") },
        { json: "contentUrl", js: "contentUrl", typ: u(undefined, "") },
        { json: "context", js: "context", typ: u(undefined, a(r("StaticTabContext"))) },
        { json: "entityId", js: "entityId", typ: "" },
        { json: "name", js: "name", typ: u(undefined, "") },
        { json: "scopes", js: "scopes", typ: a(r("CommandListScope")) },
        { json: "searchUrl", js: "searchUrl", typ: u(undefined, "") },
        { json: "websiteUrl", js: "websiteUrl", typ: u(undefined, "") },
    ], false),
    "SubscriptionOffer": o([
        { json: "offerId", js: "offerId", typ: "" },
    ], false),
    "WebApplicationInfo": o([
        { json: "id", js: "id", typ: "" },
        { json: "resource", js: "resource", typ: u(undefined, "") },
    ], false),
    "ResourceSpecificType": [
        "Application",
        "Delegated",
    ],
    "CommandListScope": [
        "groupChat",
        "personal",
        "team",
    ],
    "AuthType": [
        "apiSecretServiceAuth",
        "microsoftEntra",
        "none",
    ],
    "CommandContext": [
        "commandBox",
        "compose",
        "message",
    ],
    "InputType": [
        "choiceset",
        "date",
        "number",
        "text",
        "textarea",
        "time",
        "toggle",
    ],
    "CommandType": [
        "action",
        "query",
    ],
    "ComposeExtensionType": [
        "apiBased",
        "botBased",
    ],
    "MessageHandlerType": [
        "link",
    ],
    "ConfigurableProperty": [
        "accentColor",
        "developerUrl",
        "largeImageUrl",
        "longDescription",
        "name",
        "privacyUrl",
        "shortDescription",
        "smallImageUrl",
        "termsOfUseUrl",
    ],
    "ConfigurableTabContext": [
        "channelTab",
        "meetingChatTab",
        "meetingDetailsTab",
        "meetingSidePanel",
        "meetingStage",
        "personalTab",
        "privateChatTab",
    ],
    "MeetingSurface": [
        "sidePanel",
        "stage",
    ],
    "ConfigurableTabScope": [
        "groupChat",
        "team",
    ],
    "SupportedSharePointHost": [
        "sharePointFullPage",
        "sharePointWebPart",
    ],
    "ConnectorScope": [
        "team",
    ],
    "SourceType": [
        "bot",
    ],
    "DefaultSize": [
        "large",
        "medium",
    ],
    "Groupchat": [
        "bot",
        "connector",
        "tab",
    ],
    "DefaultInstallScope": [
        "groupChat",
        "meetings",
        "personal",
        "team",
    ],
    "DevicePermission": [
        "geolocation",
        "midi",
        "media",
        "notifications",
        "openExternal",
    ],
    "FormFactor": [
        "desktop",
        "mobile",
    ],
    "RequirementsScope": [
        "document",
        "mail",
        "presentation",
        "workbook",
    ],
    "SendMode": [
        "block",
        "promptUser",
        "softBlock",
    ],
    "ExtensionContext": [
        "default",
        "logEventMeetingDetailsAttendee",
        "mailCompose",
        "mailRead",
        "meetingDetailsAttendee",
        "meetingDetailsOrganizer",
        "onlineMeetingDetailsOrganizer",
    ],
    "PurpleType": [
        "mobileButton",
    ],
    "ItemType": [
        "menuItem",
    ],
    "FluffyType": [
        "button",
        "menu",
    ],
    "Align": [
        "after",
        "before",
    ],
    "ActionType": [
        "executeFunction",
        "openPage",
    ],
    "Lifetime": [
        "long",
        "short",
    ],
    "RuntimeType": [
        "general",
    ],
    "ManifestVersion": [
        "1.19",
    ],
    "Permission": [
        "identity",
        "messageTeamMembers",
    ],
    "StaticTabContext": [
        "channelTab",
        "meetingChatTab",
        "meetingDetailsTab",
        "meetingSidePanel",
        "meetingStage",
        "personalTab",
        "privateChatTab",
        "teamLevelApp",
    ],
    "SupportedChannelType": [
        "privateChannels",
        "sharedChannels",
    ],
};
