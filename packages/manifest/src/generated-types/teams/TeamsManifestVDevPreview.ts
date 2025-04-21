// To parse this data:
//
//   import { Convert, TeamsManifestVDevPreview } from "./file";
//
//   const teamsManifestVDevPreview = Convert.toTeamsManifestVDevPreview(json);
//
// These functions will throw an error if the JSON doesn't
// match the expected interface, even if the JSON is valid.

export interface TeamsManifestVDevPreview {
    $schema?: string;
    /**
     * The version of the schema this manifest is using.
     */
    manifestVersion: ManifestVersion;
    /**
     * The version of the app. Changes to your manifest should cause a version change. This
     * version string must follow the semver standard (http://semver.org).
     */
    version: string;
    /**
     * A unique identifier for this app. This id must be a GUID.
     */
    id: string;
    /**
     * A unique identifier for this app in reverse domain notation. E.g: com.example.myapp
     */
    packageName?:      string;
    localizationInfo?: LocalizationInfo;
    developer:         Developer;
    name:              NameClass;
    description:       Description;
    icons:             Icons;
    /**
     * A color to use in conjunction with the icon. The value must be a valid HTML color code
     * starting with '#', for example `#4464ee`.
     */
    accentColor: string;
    /**
     * These are tabs users can optionally add to their channels and 1:1 or group chats and
     * require extra configuration before they are added. Configurable tabs are not supported in
     * the personal scope. Currently only one configurable tab per app is supported.
     */
    configurableTabs?: ConfigurableTab[];
    /**
     * A set of tabs that may be 'pinned' by default, without the user adding them manually.
     * Static tabs declared in personal scope are always pinned to the app's personal
     * experience. Static tabs do not currently support the 'teams' scope.
     */
    staticTabs?: StaticTab[];
    /**
     * The set of bots for this app. Currently only one bot per app is supported.
     */
    bots?: Bot[];
    /**
     * The set of Office365 connectors for this app. Currently only one connector per app is
     * supported.
     */
    connectors?: Connector[];
    /**
     * Subscription offer associated with this app.
     */
    subscriptionOffer?: SubscriptionOffer;
    /**
     * The set of compose extensions for this app. Currently only one compose extension per app
     * is supported.
     */
    composeExtensions?: ComposeExtension[];
    scopeConstraints?:  ScopeConstraints;
    /**
     * Specifies the permissions the app requests from users.
     */
    permissions?: Permission[];
    /**
     * Specify the native features on a user's device that your app may request access to.
     */
    devicePermissions?: DevicePermission[];
    /**
     * A list of valid domains from which the tabs expect to load any content. Domain listings
     * can include wildcards, for example `*.example.com`. If your tab configuration or content
     * UI needs to navigate to any other domain besides the one use for tab configuration, that
     * domain must be specified here.
     */
    validDomains?: string[];
    /**
     * Specify your AAD App ID and Graph information to help users seamlessly sign into your AAD
     * app.
     */
    webApplicationInfo?: WebApplicationInfo;
    /**
     * Specify the app's Graph connector configuration. If this is present then
     * webApplicationInfo.id must also be specified.
     */
    graphConnector?: GraphConnector;
    /**
     * A value indicating whether or not show loading indicator when app/tab is loading
     */
    showLoadingIndicator?: boolean;
    /**
     * A value indicating whether a personal app is rendered without a tab header-bar
     */
    isFullScreen?: boolean;
    activities?:   Activities;
    /**
     * The set of supported channel type that an app belongs to
     */
    supportedChannelTypes?: SupportedChannelType[];
    /**
     * A list of tenant configured properties for an app
     */
    configurableProperties?: ConfigurableProperty[];
    /**
     * A value indicating whether an app is blocked by default until admin allows it
     */
    defaultBlockUntilAdminAction?: boolean;
    /**
     * The url to the page that provides additional app information for the admins
     */
    publisherDocsUrl?: string;
    /**
     * The install scope defined for this app by default. This will be the option displayed on
     * the button when a user tries to add the app
     */
    defaultInstallScope?: DefaultInstallScope;
    /**
     * When a group install scope is selected, this will define the default capability when the
     * user installs the app
     */
    defaultGroupCapability?: DefaultGroupCapability;
    /**
     * Specify meeting extension definition.
     */
    meetingExtensionDefinition?: MeetingExtensionDefinition;
    /**
     * Specify and consolidates authorization related information for the App.
     */
    authorization?: TeamsManifestVDevPreviewAuthorization;
    extensions?:    ElementExtension[];
    actions?:       ElementAction[];
    /**
     * Defines the list of cards which could be pinned to dashboards that can provide summarized
     * view of information relevant to user.
     */
    dashboardCards?: DashboardCard[];
    /**
     * The Intune-related properties for the app.
     */
    intuneInfo?:             IntuneInfo;
    copilotAgents?:          CopilotAgents;
    elementRelationshipSet?: ElementRelationshipSet;
    /**
     * Optional property containing background loading configuration. By opting into this
     * performance enhancement, your app is eligible to be loaded in the background in any
     * Microsoft 365 application host that supports this feature. Note that setting this
     * property gives the host client permission to load the app in the background but does not
     * guarantee the app will be preloaded at runtime. Whether an app is preloaded in the
     * background will be dynamically determined based on usage and other criteria.
     */
    backgroundLoadConfiguration?: BackgroundLoadConfiguration;
}

/**
 * Actions node contains an array of actions object.
 */
export interface ElementAction {
    /**
     * A unique identifier string in the default locale that is used to catalog actions.
     */
    id: string;
    /**
     * An enum string that describes the intent of the action.
     */
    intent: Intent;
    /**
     * A display name for the action.
     */
    displayName: string;
    /**
     * A display string in the default locale to represent the action.
     */
    description: string;
    /**
     * Object containing URLs to icon images for this action intent.
     */
    icons?: Icon[];
    /**
     * Defining how actions can be handled. If an app has more than 1 handler, only one
     * experience will show up at one entry point. The hub will decide which action to show up
     * based on which experience is supported.
     */
    handlers: Handler[];
}

export interface Handler {
    /**
     * Required both for File Handlers and Content Actions.
     */
    type:              HandlerType;
    supportedObjects?: SupportedObjects;
    /**
     * If true, multiple files can be selected and the action will still be displayed. If false
     * or missing, the action is only displayed when a single item is selected.
     */
    supportsMultiSelect?: boolean;
    pageInfo?:            PageInfo;
    dialogInfo?:          DialogInfo;
    /**
     * Url for handler type openURL, invokeAPI, openTaskpane, and others.
     */
    url?:     string;
    botInfo?: BotInfo;
    [property: string]: any;
}

export interface BotInfo {
    /**
     * Bot ID.
     */
    botId: string;
    /**
     * Fetch task from bot.
     */
    fetchTask?: boolean;
    [property: string]: any;
}

export interface DialogInfo {
    /**
     * Dialog type, defines how the developer build the dialog.
     */
    dialogType: DialogType;
    /**
     * Required for html based dialog.
     */
    url?: string;
    /**
     * Dialog width - either a number in pixels or default layout such as 'large', 'medium', or
     * 'small'.
     */
    width: string;
    /**
     * Dialog height - either a number in pixels or default layout such as 'large', 'medium', or
     * 'small'.
     */
    height: string;
    /**
     * Array of parameter object, each contains: name, title, description, inputType.
     */
    parameters?: ParameterObject[];
    /**
     * Dialog title.
     */
    title?: string;
    [property: string]: any;
}

/**
 * Dialog type, defines how the developer build the dialog.
 */
export type DialogType = "url" | "adaptiveCard";

export interface ParameterObject {
    /**
     * Parameter name.
     */
    name: string;
    /**
     * Parameter title.
     */
    title: string;
    /**
     * Parameter description.
     */
    description: string;
    /**
     * Parameter input type.
     */
    inputType: string;
    [property: string]: any;
}

export interface PageInfo {
    /**
     * Used to navigate to the page in MetaOS app.
     */
    pageId: string;
    /**
     * Used to navigate to the subpage in MetaOS app.
     */
    subpageId?: string;
    [property: string]: any;
}

export interface SupportedObjects {
    file?: File;
    /**
     * A null value indicates that the file handler is not available when a folder is selected.
     * An object with no parameters indicates that the file handler is available when a folder
     * is selected or when no files are selected.
     */
    folder?: { [key: string]: any } | null;
    [property: string]: any;
}

export interface File {
    extensions?: string[];
}

/**
 * Required both for File Handlers and Content Actions.
 */
export type HandlerType = "openURL" | "openPage" | "openDialog" | "openTaskpane" | "invokeAPI" | "invokeBot";

export interface Icon {
    /**
     * Icon size in pixels.
     */
    size: number;
    /**
     * URL for the icon.
     */
    url: string;
}

/**
 * An enum string that describes the intent of the action.
 */
export type Intent = "create" | "addTo" | "open" | "preview" | "share" | "sign" | "custom";

export interface Activities {
    /**
     * Specify the types of activites that your app can post to a users activity feed
     */
    activityTypes?: ActivityType[];
    /**
     * Specify the customized icons that your app can post to a users activity feed
     */
    activityIcons?: ActivityIcon[];
}

export interface ActivityIcon {
    /**
     * Represents the unique icon ID.
     */
    id: string;
    /**
     * Represents the relative path to the icon image. Image should be size 32x32.
     */
    iconFile: string;
}

export interface ActivityType {
    type:         string;
    description:  string;
    templateText: string;
    /**
     * An array containing valid icon IDs per activity type.
     */
    allowedIconIds?: string[];
}

/**
 * Specify and consolidates authorization related information for the App.
 */
export interface TeamsManifestVDevPreviewAuthorization {
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
     * Permissions that guard data access on a resource instance level.
     */
    resourceSpecific?: ResourceSpecific[];
}

export interface ResourceSpecific {
    /**
     * The name of the resource-specific permission.
     */
    name: string;
    /**
     * The type of the resource-specific permission.
     */
    type: ResourceSpecificType;
}

/**
 * The type of the resource-specific permission.
 */
export type ResourceSpecificType = "Application" | "Delegated";

/**
 * Optional property containing background loading configuration. By opting into this
 * performance enhancement, your app is eligible to be loaded in the background in any
 * Microsoft 365 application host that supports this feature. Note that setting this
 * property gives the host client permission to load the app in the background but does not
 * guarantee the app will be preloaded at runtime. Whether an app is preloaded in the
 * background will be dynamically determined based on usage and other criteria.
 */
export interface BackgroundLoadConfiguration {
    /**
     * Optional property within backgroundLoadConfiguration containing tab settings for
     * background loading. Setting tabConfiguration indicates that the app supports background
     * loading of tabs.
     */
    tabConfiguration?: TabConfiguration;
}

/**
 * Optional property within backgroundLoadConfiguration containing tab settings for
 * background loading. Setting tabConfiguration indicates that the app supports background
 * loading of tabs.
 */
export interface TabConfiguration {
    /**
     * Required URL for background loading. This can be the same contentUrl from the staticTabs
     * section or an alternative endpoint used for background loading.
     */
    contentUrl: string;
}

export interface Bot {
    /**
     * The Microsoft App ID specified for the bot in the Bot Framework portal
     * (https://dev.botframework.com/bots)
     */
    botId:          string;
    configuration?: Configuration;
    /**
     * Optional Boolean property that enables users to delete messages sent by bot when set to
     * true. The default value is false, ensuring bot messages cannot be deleted by users unless
     * explicitly opted in.
     */
    allowBotMessageDeleteByUser?: boolean;
    /**
     * This value describes whether or not the bot utilizes a user hint to add the bot to a
     * specific channel.
     */
    needsChannelSelector?: boolean;
    /**
     * A value indicating whether or not the bot is a one-way notification only bot, as opposed
     * to a conversational bot.
     */
    isNotificationOnly?: boolean;
    /**
     * A value indicating whether the team's Office group needs to be security enabled.
     */
    requiresSecurityEnabledGroup?: boolean;
    /**
     * A value indicating whether the bot supports uploading/downloading of files.
     */
    supportsFiles?: boolean;
    /**
     * A value indicating whether the bot supports audio calling.
     */
    supportsCalling?: boolean;
    /**
     * A value indicating whether the bot supports video calling.
     */
    supportsVideo?: boolean;
    /**
     * Specifies whether the bot offers an experience in the context of a channel in a team, in
     * a group chat (groupChat), an experience scoped to an individual user alone (personal) OR
     * within Copilot surfaces. These options are non-exclusive.
     */
    scopes: CommandListScope[];
    /**
     * The list of commands that the bot supplies, including their usage, description, and the
     * scope for which the commands are valid. A separate command list should be used for each
     * scope.
     */
    commandLists?: CommandList[];
    /**
     * The set of requirements for the bot.
     */
    requirementSet?: ElementRequirementSet;
}

export interface CommandList {
    /**
     * Specifies whether the bot offers an experience in the context of a channel in a team, in
     * a group chat (groupChat), an experience scoped to an individual user alone (personal) OR
     * within Copilot surfaces. These options are non-exclusive.
     */
    scopes:   CommandListScope[];
    commands: CommandListCommand[];
}

export interface CommandListCommand {
    /**
     * The bot command name
     */
    title: string;
    /**
     * A simple text description or an example of the command syntax and its arguments.
     */
    description: string;
}

export type CommandListScope = "team" | "personal" | "groupChat" | "copilot";

export interface Configuration {
    team?:      ConfigurationTeam;
    groupChat?: ConfigurationGroupChat;
}

export interface ConfigurationGroupChat {
    /**
     * A boolean value that indicates if it should fetch bot config task module dynamically.
     */
    fetchTask?: boolean;
    /**
     * Task module to be launched when fetch task set to false.
     */
    taskInfo?: TaskInfo;
}

/**
 * Task module to be launched when fetch task set to false.
 */
export interface TaskInfo {
    /**
     * Initial dialog title.
     */
    title?: string;
    /**
     * Dialog width - either a number in pixels or default layout such as 'large', 'medium', or
     * 'small'.
     */
    width?: string;
    /**
     * Dialog height - either a number in pixels or default layout such as 'large', 'medium', or
     * 'small'.
     */
    height?: string;
    /**
     * Initial webview URL.
     */
    url?: string;
}

export interface ConfigurationTeam {
    /**
     * A boolean value that indicates if it should fetch bot config task module dynamically.
     */
    fetchTask?: boolean;
    /**
     * Task module to be launched when fetch task set to false.
     */
    taskInfo?: TaskInfo;
}

/**
 * The set of requirements for the tab.
 *
 * An object representing a set of requirements that the host must support for the element.
 *
 * The set of requirements for the bot.
 *
 * The set of requirements for the compose extension.
 */
export interface ElementRequirementSet {
    hostMustSupportFunctionalities: HostFunctionality[];
}

/**
 * An object representing a specific functionality that a host must support.
 */
export interface HostFunctionality {
    /**
     * The name of the functionality.
     */
    name: HostMustSupportFunctionalityName;
}

/**
 * The name of the functionality.
 */
export type HostMustSupportFunctionalityName = "dialogUrl" | "dialogUrlBot" | "dialogAdaptiveCard" | "dialogAdaptiveCardBot";

export interface ComposeExtension {
    /**
     * A unique identifier for the compose extension.
     */
    id?: string;
    /**
     * The Microsoft App ID specified for the bot powering the compose extension in the Bot
     * Framework portal (https://dev.botframework.com/bots)
     */
    botId?: string;
    /**
     * Type of the compose extension.
     */
    composeExtensionType?: ComposeExtensionType;
    /**
     * Object capturing authorization information.
     */
    authorization?: ComposeExtensionAuthorization;
    /**
     * A relative file path to the api specification file in the manifest package.
     */
    apiSpecificationFile?: string;
    /**
     * A value indicating whether the configuration of a compose extension can be updated by the
     * user.
     */
    canUpdateConfiguration?: boolean | null;
    commands?:               ComposeExtensionCommand[];
    /**
     * A list of handlers that allow apps to be invoked when certain conditions are met
     */
    messageHandlers?: MessageHandler[];
    /**
     * The set of requirements for the compose extension.
     */
    requirementSet?: ElementRequirementSet;
}

/**
 * Object capturing authorization information.
 */
export interface ComposeExtensionAuthorization {
    /**
     * Enum of possible authorization types.
     */
    authType?: AuthType;
    /**
     * Object capturing details needed to do microsoftEntra auth flow. It will be only present
     * when auth type is microsoftEntra.
     */
    microsoftEntraConfiguration?: MicrosoftEntraConfiguration;
    /**
     * Object capturing details needed to do service auth. It will be only present when auth
     * type is apiSecretServiceAuth.
     */
    apiSecretServiceAuthConfiguration?: APISecretServiceAuthConfiguration;
    /**
     * Object capturing details needed to match the application's OAuth configuration for the
     * app. This should be and must be populated only when authType is set to oAuth2.0r
     */
    oAuthConfiguration?: OAuthConfiguration;
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
 * Enum of possible authorization types.
 */
export type AuthType = "none" | "apiSecretServiceAuth" | "microsoftEntra" | "oAuth2.0";

/**
 * Object capturing details needed to do microsoftEntra auth flow. It will be only present
 * when auth type is microsoftEntra.
 */
export interface MicrosoftEntraConfiguration {
    /**
     * Boolean indicating whether single sign on is configured for the app.
     */
    supportsSingleSignOn?: boolean;
}

/**
 * Object capturing details needed to match the application's OAuth configuration for the
 * app. This should be and must be populated only when authType is set to oAuth2.0r
 */
export interface OAuthConfiguration {
    /**
     * The oAuth configuration id obtained by the Developer when registering their configuration
     * in Developer Portal.
     */
    oAuthConfigurationId?: string;
}

export interface ComposeExtensionCommand {
    /**
     * Id of the command.
     */
    id: string;
    /**
     * Type of the command
     */
    type?:          CommandType;
    samplePrompts?: SamplePrompt[];
    /**
     * A relative file path for api response rendering template file. The schema of the file can
     * be referred to in this
     * link:'https://developer.microsoft.com/json-schemas/teams/vDevPreview/MicrosoftTeams.ResponseRenderingTemplate.schema.json'.
     */
    apiResponseRenderingTemplateFile?: string;
    /**
     * Context where the command would apply
     */
    context?: CommandContext[];
    /**
     * Title of the command.
     */
    title: string;
    /**
     * Description of the command.
     */
    description?: string;
    /**
     * A boolean value that indicates if the command should be run once initially with no
     * parameter.
     */
    initialRun?: boolean;
    /**
     * A boolean value that indicates if it should fetch task module dynamically
     */
    fetchTask?:  boolean;
    parameters?: ParameterClass[];
    /**
     * Task module to be launched when fetch task set to false.
     */
    taskInfo?: TaskInfo;
    /**
     * semantic description of the command. This is typically meant for consumption by the large
     * language model.
     */
    semanticDescription?: string;
}

export type CommandContext = "compose" | "commandBox" | "message";

export interface ParameterClass {
    /**
     * Name of the parameter.
     */
    name: string;
    /**
     * Type of the parameter
     */
    inputType?: InputType;
    /**
     * Indicates whether this parameter is required or not. By default, it is not.
     */
    isRequired?: boolean;
    /**
     * Title of the parameter.
     */
    title: string;
    /**
     * Description of the parameter.
     */
    description?: string;
    /**
     * Initial value for the parameter
     */
    value?: string;
    /**
     * The choice options for the parameter
     */
    choices?: Choice[];
    /**
     * semantic description of the parameter. This is typically meant for consumption by the
     * large language model.
     */
    semanticDescription?: string;
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
export type InputType = "text" | "textarea" | "number" | "date" | "time" | "toggle" | "choiceset";

export interface SamplePrompt {
    /**
     * This string will hold the sample prompt
     */
    text: string;
}

/**
 * Type of the command
 */
export type CommandType = "query" | "action";

/**
 * Type of the compose extension.
 */
export type ComposeExtensionType = "botBased" | "apiBased";

export interface MessageHandler {
    /**
     * Type of the message handler
     */
    type:  "link";
    value: Value;
}

/**
 * Type of the message handler
 */

export interface Value {
    /**
     * A list of domains that the link message handler can register for, and when they are
     * matched the app will be invoked
     */
    domains?: string[];
    /**
     * A boolean value that indicates whether the app's link message handler supports anonymous
     * invoke flow. [Deprecated]. This property has been superceded by
     * 'supportsAnonymizedPayloads'.
     */
    supportsAnonymousAccess?: boolean;
    /**
     * A boolean value that indicates whether the app's link message handler supports anonymous
     * invoke flow.
     */
    supportsAnonymizedPayloads?: boolean;
    [property: string]: any;
}

export type ConfigurableProperty = "name" | "shortDescription" | "longDescription" | "smallImageUrl" | "largeImageUrl" | "accentColor" | "developerUrl" | "privacyUrl" | "termsOfUseUrl";

export interface ConfigurableTab {
    /**
     * A unique identifier for the tab. This id must be unique within the app manifest.
     */
    id?: string;
    /**
     * The url to use when configuring the tab.
     */
    configurationUrl: string;
    /**
     * A value indicating whether an instance of the tab's configuration can be updated by the
     * user after creation.
     */
    canUpdateConfiguration?: boolean;
    /**
     * Specifies whether the tab offers an experience in the context of a channel in a team, in
     * a 1:1 or group chat, or in an experience scoped to an individual user alone. These
     * options are non-exclusive. Currently, configurable tabs are only supported in the teams
     * and groupchats scopes.
     */
    scopes: ConfigurableTabScope[];
    /**
     * The set of meetingSurfaceItem scopes that a tab belong to
     */
    meetingSurfaces?: MeetingSurface[];
    /**
     * The set of contextItem scopes that a tab belong to
     */
    context?: ConfigurableTabContext[];
    /**
     * The set of supportedPlatform scopes that a tab belong to
     */
    supportedPlatform?: SupportedPlatform[];
    /**
     * A relative file path to a tab preview image for use in SharePoint. Size 1024x768.
     */
    sharePointPreviewImage?: string;
    /**
     * Defines how your tab will be made available in SharePoint.
     */
    supportedSharePointHosts?: SupportedSharePointHost[];
}

export type ConfigurableTabContext = "personalTab" | "channelTab" | "privateChatTab" | "meetingChatTab" | "meetingDetailsTab" | "meetingSidePanel" | "meetingStage" | "callingSidePanel";

export type MeetingSurface = "sidePanel" | "stage";

export type ConfigurableTabScope = "team" | "groupChat";

export type SupportedPlatform = "desktop" | "mobile" | "teamsMeetingDevices";

export type SupportedSharePointHost = "sharePointFullPage" | "sharePointWebPart";

export interface Connector {
    /**
     * A unique identifier for the connector which matches its ID in the Connectors Developer
     * Portal.
     */
    connectorId: string;
    /**
     * The url to use for configuring the connector using the inline configuration experience.
     */
    configurationUrl?: string;
    /**
     * Specifies whether the connector offers an experience in the context of a channel in a
     * team, or an experience scoped to an individual user alone. Currently, only the team scope
     * is supported.
     */
    scopes: "team"[];
}

export interface CopilotAgents {
    /**
     * An array of declarative agent elements references. Currently, only one declarative agent
     * per application is supported.
     */
    declarativeAgents?: DeclarativeAgentRef[];
    /**
     * An array of Custom Engine Agents. Currently only one Custom Engine Agent per application
     * is supported.
     */
    customEngineAgents?: CustomEngineAgent[];
}

export interface CustomEngineAgent {
    /**
     * The id of the Custom Engine Agent. If it is of type bot, the id must match the id
     * specified in a bot in the bots node and the referenced bot must have personal scope. The
     * app short name and short description must also be defined.
     */
    id: string;
    /**
     * The type of the Custom Engine Agent. Currently only type bot is supported.
     */
    type: "bot";
}

/**
 * The type of the Custom Engine Agent. Currently only type bot is supported.
 *
 * The content of the dashboard card is sourced from a bot.
 */

/**
 * A reference to a declarative agent element. The element's definition is in a separate
 * file.
 */
export interface DeclarativeAgentRef {
    /**
     * A unique identifier for this declarative agent element.
     */
    id: string;
    /**
     * Relative file path to this declarative agent element file in the application package.
     */
    file: string;
}

/**
 * Cards wich could be pinned to dashboard providing summarized view of information relevant
 * to user.
 */
export interface DashboardCard {
    /**
     * Unique Id for the card. Must be unique inside the app.
     */
    id: string;
    /**
     * Represents the name of the card. Maximum length is 255 characters.
     */
    displayName: string;
    /**
     * Description of the card.Maximum length is 255 characters.
     */
    description: string;
    /**
     * Id of the group in the card picker. This must be guid.
     */
    pickerGroupId: string;
    icon?:         DashboardCardIcon;
    contentSource: DashboardCardContentSource;
    /**
     * Rendering Size for dashboard card.
     */
    defaultSize: DefaultSize;
}

/**
 * Represents a configuration for the source of the card’s content.
 */
export interface DashboardCardContentSource {
    /**
     * The content of the dashboard card is sourced from a bot.
     */
    sourceType?: "bot";
    /**
     * The configuration for the bot source. Required if sourceType is set to bot.
     */
    botConfiguration?: BotConfiguration;
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
 * Rendering Size for dashboard card.
 */
export type DefaultSize = "medium" | "large";

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
     * When the install scope selected is Team, this field specifies the default capability
     * available
     */
    team?: Groupchat;
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
export type Groupchat = "tab" | "bot" | "connector";

/**
 * The install scope defined for this app by default. This will be the option displayed on
 * the button when a user tries to add the app
 */
export type DefaultInstallScope = "personal" | "team" | "groupChat" | "meetings" | "copilot";

export interface Description {
    /**
     * A short description of the app used when space is limited. Maximum length is 80
     * characters.
     */
    short: string;
    /**
     * The full description of the app. Maximum length is 4000 characters.
     */
    full: string;
}

export interface Developer {
    /**
     * The display name for the developer.
     */
    name: string;
    /**
     * The Microsoft Partner Network ID that identifies the partner organization building the
     * app. This field is not required, and should only be used if you are already part of the
     * Microsoft Partner Network. More info at https://aka.ms/partner
     */
    mpnId?: string;
    /**
     * The url to the page that provides support information for the app.
     */
    websiteUrl: string;
    /**
     * The url to the page that provides privacy information for the app.
     */
    privacyUrl: string;
    /**
     * The url to the page that provides the terms of use for the app.
     */
    termsOfUseUrl: string;
    /**
     * App developer contact information.
     */
    contactInfo?: ContactInfo;
}

/**
 * App developer contact information.
 */
export interface ContactInfo {
    /**
     * Support configuration.
     */
    defaultSupport: DefaultSupport;
    [property: string]: any;
}

/**
 * Support configuration.
 */
export interface DefaultSupport {
    /**
     * User email for chat support contacts.
     */
    userEmailsForChatSupport: string[];
    /**
     * Email address for email support.
     */
    emailsForEmailSupport: string[];
    [property: string]: any;
}

export type DevicePermission = "geolocation" | "media" | "notifications" | "midi" | "openExternal";

export interface ElementRelationshipSet {
    /**
     * An array containing multiple instances of unidirectional dependency relationships (each
     * represented by a oneWayDependency object).
     */
    oneWayDependencies?: OneWayDependency[];
    /**
     * An array containing multiple instances of mutual dependency relationships between
     * elements (each represented by a mutualDependency object).
     */
    mutualDependencies?: Array<ElementReference[]>;
}

/**
 * A specific instance of mutual dependency between two or more elements, indicating that
 * each element depends on the others in a bidirectional manner.
 */
export interface ElementReference {
    name:        MutualDependencyName;
    id:          string;
    commandIds?: string[];
}

export type MutualDependencyName = "bots" | "staticTabs" | "composeExtensions" | "configurableTabs";

/**
 * An object representing a unidirectional dependency relationship, where one specific
 * element (referred to as the `element`) relies on an array of other elements (referred to
 * as the `dependsOn`) in a single direction.
 */
export interface OneWayDependency {
    element:   ElementReference;
    dependsOn: ElementReference[];
}

/**
 * The set of extensions for this app. Currently only one extensions per app is supported.
 */
export interface ElementExtension {
    requirements?: RequirementsExtensionElement;
    /**
     * General runtime for "MailApp" or "TaskpaneApp". Configures the set of runtimes and
     * actions that can be used by each extension point. Min size 1.
     */
    runtimes?:      ExtensionRuntimesArray[];
    ribbons?:       ExtensionRibbonsArray[];
    autoRunEvents?: ExtensionAutoRunEventsArray[];
    alternates?:    ExtensionAlternateVersionsArray[];
    /**
     * The url for your extension, used to validate Exchange user identity tokens.
     */
    audienceClaimUrl?:   string;
    appDeeplinks?:       ExtensionAppDeeplinksArray[];
    contentRuntimes?:    ExtensionContentRuntimeArray[];
    getStartedMessages?: ExtensionGetStartedMessageArray[];
    /**
     * Specifies the context menus for your extension. A context menu is a shortcut menu that
     * appears when a user right-clicks (selects and holds) in the Office UI. Min size 1.
     */
    contextMenus?: ExtensionContextMenuArray[];
    /**
     * Keyboard shortcuts, also known as key combinations, enable your add-in's users to work
     * more efficiently. Keyboard shortcuts also improve the add-in's accessibility for users
     * with disabilities by providing an alternative to the mouse.
     */
    keyboardShortcuts?: ExtensionKeyboardShortcut[];
}

export interface ExtensionAlternateVersionsArray {
    requirements?:       RequirementsExtensionElement;
    prefer?:             Prefer;
    hide?:               Hide;
    alternateIcons:      AlternateIcons;
    xllCustomFunctions?: ExtensionXllCustomFunctions;
}

export interface AlternateIcons {
    icon:               ExtensionCommonIcon;
    highResolutionIcon: ExtensionCommonIcon;
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
    storeOfficeAddin?:  StoreOfficeAddin;
    customOfficeAddin?: CustomOfficeAddin;
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
     * Solution ID of an in-market add-in to hide. Maximum length is 64 characters.
     */
    officeAddinId: string;
    /**
     * Asset ID of the in-market add-in to hide. Maximum length is 64 characters.
     */
    assetId: string;
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

/**
 * Specifies limitations on which clients the add-in can be installed on, including
 * limitations on the Office host application, the form factors, and the requirement sets
 * that the client must support.
 *
 * Specifies the Office requirement sets.
 */
export interface RequirementsExtensionElement {
    capabilities?: Capability[];
    /**
     * Identifies the scopes in which the add-in can run. For example, mail means Outlook.
     */
    scopes?: RequirementsScope[];
    /**
     * Identifies the form factors that support the add-in. Supported values: mobile, desktop.
     */
    formFactors?: FormFactor[];
}

export interface Capability {
    /**
     * Identifies the name of the requirement sets that the add-in needs to run.
     */
    name: string;
    /**
     * Identifies the minimum version for the requirement sets that the add-in needs to run.
     */
    minVersion?: string;
    /**
     * Identifies the maximum version for the requirement sets that the add-in needs to run.
     */
    maxVersion?: string;
}

export type FormFactor = "desktop" | "mobile";

export type RequirementsScope = "mail" | "workbook" | "document" | "presentation";

export interface ExtensionXllCustomFunctions {
    fileName?: string;
    [property: string]: any;
}

/**
 * Represents the copilot extension point
 */
export interface ExtensionAppDeeplinksArray {
    requirements?: AppDeeplinkRequirements;
    contexts:      ExtensionContext[];
    /**
     * The ID of an action defined in runtimes. Manifest should be invalidated if no action with
     * an id matching actionId is present in runtimes.
     */
    actionId: string;
    /**
     * the text that will be shown on the app as a clickable item.
     */
    label: string;
    /**
     * the text metadata, for recommendation engine.
     */
    semanticDescription: string;
}

/**
 * Specifies the Office application windows in which the ribbon customization is available
 * to the user. Each item in the array is a member of a string array. Possible values are:
 * mailRead, mailCompose, meetingDetailsOrganizer, meetingDetailsAttendee,
 * onlineMeetingDetailsOrganizer, logEventMeetingDetailsAttendee, spamReportingOverride.
 */
export type ExtensionContext = "mailRead" | "mailCompose" | "meetingDetailsOrganizer" | "meetingDetailsAttendee" | "onlineMeetingDetailsOrganizer" | "logEventMeetingDetailsAttendee" | "default" | "spamReportingOverride";

/**
 * Specifies limitations on which clients the add-in can be installed on, including
 * limitations on the Office host application, the form factors, and the requirement sets
 * that the client must support.
 *
 * Specifies the Office requirement sets.
 */
export interface AppDeeplinkRequirements {
    capabilities?: Capability[];
    /**
     * Identifies the scopes in which the add-in can run. For example, mail means Outlook.
     */
    scopes?: RequirementsScope[];
    /**
     * Identifies the form factors that support the add-in. Supported values: mobile, desktop.
     */
    formFactors?: FormFactor[];
}

export interface ExtensionAutoRunEventsArray {
    requirements?: RequirementsExtensionElement;
    /**
     * Specifies the type of event. For supported types, please see:
     * https://learn.microsoft.com/en-us/office/dev/add-ins/outlook/autolaunch?tabs=xmlmanifest#supported-events.
     */
    events: Event[];
}

export interface Event {
    type: string;
    /**
     * The ID of an action defined in runtimes. Maximum length is 64 characters.
     */
    actionId: string;
    /**
     * Configures how Outlook responds to the event.
     */
    options?: Options;
}

/**
 * Configures how Outlook responds to the event.
 */
export interface Options {
    sendMode: SendMode;
}

export type SendMode = "promptUser" | "softBlock" | "block";

/**
 * Content runtime is for 'ContentApp', which can be embedded directly into Excel or
 * PowerPoint documents.
 */
export interface ExtensionContentRuntimeArray {
    requirements?: ContentRuntimeRequirements;
    /**
     * A unique identifier for this runtime within the app. This is developer specified.
     */
    id:   string;
    code: ExtensionRuntimeCode;
    /**
     * The desired height in pixels for the initial content placeholder. This value MUST be
     * between 32 and 1000 pixels. Default value will be determined by host.
     */
    requestedHeight?: number;
    /**
     * The desired width in pixels for the initial content placeholder. This value MUST be
     * between 32 and 1000 pixels. Default value will be determined by host.
     */
    requestedWidth?: number;
    /**
     * Specifies whether a snapshot image of your content add-in is saved with the host
     * document. Default value is false. Set true to disable.
     */
    disableSnapshot?: boolean;
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
 * Specifies limitations on which clients the add-in can be installed on, including
 * limitations on the Office host application, the form factors, and the requirement sets
 * that the client must support.
 *
 * Specifies the Office requirement sets.
 */
export interface ContentRuntimeRequirements {
    capabilities?: Capability[];
    /**
     * Identifies the scopes in which the add-in can run. For example, mail means Outlook.
     */
    scopes?: RequirementsScope[];
    /**
     * Identifies the form factors that support the add-in. Supported values: mobile, desktop.
     */
    formFactors?: FormFactor[];
}

/**
 * Specifies the context menus for your extension. A context menu is a shortcut menu that
 * appears when a user right-clicks (selects and holds) in the Office UI. Min size 1.
 */
export interface ExtensionContextMenuArray {
    requirements?: ContextMenuRequirements;
    /**
     * Configures the context menus. Minimum size is 1.
     */
    menus: ExtensionMenuItem[];
}

/**
 * Configures the context menus. Minimum size is 1.
 *
 * The title used for the top of the callout.
 */
export interface ExtensionMenuItem {
    /**
     * Use 'text' or 'cell' here for Office context menu. Use text if the context menu should
     * open when a user right-clicks (selects and holds) on the selected text. Use cell if the
     * context menu should open when the user right-clicks (selects and holds) on a cell in an
     * Excel spreadsheet.
     */
    entryPoint: EntryPoint;
    controls:   ExtensionCommonCustomGroupControlsItem[];
}

export interface ExtensionCommonCustomGroupControlsItem {
    /**
     * A unique identifier for this control within the app. Maximum length is 64 characters.
     */
    id: string;
    /**
     * Defines the type of control whether button or menu.
     */
    type: PurpleType;
    /**
     * Id of an existing office control. Maximum length is 64 characters.
     */
    builtInControlId?: string;
    /**
     * Displayed text for the control. Maximum length is 64 characters.
     */
    label: string;
    /**
     * Configures the icons for the custom control.
     */
    icons:    ExtensionCommonIcon[];
    supertip: ExtensionCommonSuperToolTip;
    /**
     * The ID of an execution-type action that handles this key combination. Maximum length is
     * 64 characters.
     */
    actionId?: string;
    /**
     * Specifies whether a group, button, menu, or menu item will be hidden on application and
     * platform combinations that support the API (Office.ribbon.requestCreateControls) that
     * installs custom contextual tabs on the ribbon. Default is false.
     */
    overriddenByRibbonApi?: boolean;
    /**
     * Whether the control is initially enabled.
     */
    enabled?: boolean;
    /**
     * Configures the items for a menu control.
     */
    items?: ExtensionCommonCustomControlMenuItem[];
}

export interface ExtensionCommonCustomControlMenuItem {
    /**
     * A unique identifier for this control within the app. Maximum length is 64 characters.
     */
    id: string;
    /**
     * Supported values: menuItem.
     */
    type: "menuItem";
    /**
     * Displayed text for the control. Maximum length is 64 characters.
     */
    label:    string;
    icons?:   ExtensionCommonIcon[];
    supertip: ExtensionCommonSuperToolTip;
    /**
     * The ID of an action defined in runtimes. Maximum length is 64 characters.
     */
    actionId: string;
    /**
     * Whether the control is initially enabled.
     */
    enabled?:               boolean;
    overriddenByRibbonApi?: boolean;
}

export interface ExtensionCommonSuperToolTip {
    /**
     * Title text of the super tip. Maximum length is 64 characters.
     */
    title: string;
    /**
     * Description of the super tip. Maximum length is 250 characters.
     */
    description: string;
}

/**
 * Supported values: menuItem.
 */

/**
 * Defines the type of control whether button or menu.
 */
export type PurpleType = "button" | "menu";

/**
 * Use 'text' or 'cell' here for Office context menu. Use text if the context menu should
 * open when a user right-clicks (selects and holds) on the selected text. Use cell if the
 * context menu should open when the user right-clicks (selects and holds) on a cell in an
 * Excel spreadsheet.
 */
export type EntryPoint = "text" | "cell";

/**
 * Specifies limitations on which clients the add-in can be installed on, including
 * limitations on the Office host application, the form factors, and the requirement sets
 * that the client must support.
 *
 * Specifies the Office requirement sets.
 */
export interface ContextMenuRequirements {
    capabilities?: Capability[];
    /**
     * Identifies the scopes in which the add-in can run. For example, mail means Outlook.
     */
    scopes?: RequirementsScope[];
    /**
     * Identifies the form factors that support the add-in. Supported values: mobile, desktop.
     */
    formFactors?: FormFactor[];
}

/**
 * Provides information used by the callout that appears when the add-in is installed.
 */
export interface ExtensionGetStartedMessageArray {
    requirements?: GetStartedMessageRequirements;
    /**
     * The title used for the top of the callout.
     */
    title: string;
    /**
     * The description/body content for the callout.
     */
    description: string;
    /**
     * A URL to a page that explains the add-in in detail.
     */
    learnMoreUrl: string;
}

/**
 * Specifies limitations on which clients the add-in can be installed on, including
 * limitations on the Office host application, the form factors, and the requirement sets
 * that the client must support.
 *
 * Specifies the Office requirement sets.
 */
export interface GetStartedMessageRequirements {
    capabilities?: Capability[];
    /**
     * Identifies the scopes in which the add-in can run. For example, mail means Outlook.
     */
    scopes?: RequirementsScope[];
    /**
     * Identifies the form factors that support the add-in. Supported values: mobile, desktop.
     */
    formFactors?: FormFactor[];
}

export interface ExtensionKeyboardShortcut {
    /**
     * Specifies the Office requirement sets.
     */
    requirements?: RequirementsExtensionElement;
    /**
     * Array of mappings from actions to the key combinations that invoke the actions.
     */
    shortcuts: ExtensionShortcut[];
    [property: string]: any;
}

export interface ExtensionShortcut {
    key: Key;
    /**
     * The ID of an execution-type action that handles this key combination.
     */
    actionId: string;
    [property: string]: any;
}

/**
 * Key combinations in different platform (i.e. default, windows, web and mac).
 */
export interface Key {
    /**
     * Fallback key for any platform that isn't specified.
     */
    default: string;
    /**
     * key for mac platform. Alt is mapped to the Option key.
     */
    mac?: string;
    /**
     * key for web platform.
     */
    web?: string;
    /**
     * key for windows platform. Command is mapped to the Ctrl key.
     */
    windows?: string;
    [property: string]: any;
}

export interface ExtensionRibbonsArray {
    requirements?:            RequirementsExtensionElement;
    contexts?:                ExtensionContext[];
    tabs:                     ExtensionRibbonsArrayTabsItem[];
    fixedControls?:           ExtensionRibbonsArrayFixedControlItem[];
    spamPreProcessingDialog?: ExtensionRibbonsSpamPreProcessingDialog;
}

export interface ExtensionRibbonsArrayFixedControlItem {
    /**
     * A unique identifier for this control within the app. Maximum length is 64 characters.
     */
    id: string;
    /**
     * Defines the type of control.
     */
    type: "button";
    /**
     * Displayed text for the control. Maximum length is 64 characters.
     */
    label:    string;
    icons:    ExtensionCommonIcon[];
    supertip: ExtensionCommonSuperToolTip;
    /**
     * The ID of an execution-type action that handles this key combination. Maximum length is
     * 64 characters.
     */
    actionId: string;
    /**
     * Whether the control is initially enabled.
     */
    enabled: boolean;
}

/**
 * Defines the type of control.
 */

export interface ExtensionRibbonsSpamPreProcessingDialog {
    /**
     * Specifies the custom title of the preprocessing dialog.
     */
    title: string;
    /**
     * Specifies the custom text that appears in the preprocessing dialog.
     */
    description: string;
    /**
     * Specifies whether the bot offers an experience in the context of a channel in a team, in
     * a 1:1 or group chat, or in an experience scoped to an individual user alone. These
     * options are non-exclusive.
     */
    spamReportingOptions?: SpamReportingOptions;
    /**
     * A text box to the preprocessing dialog to allow users to provide additional information
     * on the message they're reporting. This value is the title of that text box.
     */
    spamFreeTextSectionTitle?: string;
    /**
     * Specifies the custom text and URL to provide informational resources to the users.
     */
    spamMoreInfo?: SpamMoreInfo;
}

/**
 * Specifies the custom text and URL to provide informational resources to the users.
 */
export interface SpamMoreInfo {
    /**
     * Specifies display content of the hyperlink pointing to the site containing informational
     * resources in the preprocessing dialog of a spam-reporting add-in.
     */
    text: string;
    /**
     * Specifies the URL of the hyperlink pointing to the site containing informational
     * resources in the preprocessing dialog of a spam-reporting add-in.
     */
    url: string;
    [property: string]: any;
}

/**
 * Specifies whether the bot offers an experience in the context of a channel in a team, in
 * a 1:1 or group chat, or in an experience scoped to an individual user alone. These
 * options are non-exclusive.
 */
export interface SpamReportingOptions {
    /**
     * Specifies the title listed before the reporting options list.
     */
    title: string;
    /**
     * Specifies the custom options that a user can select from the preprocessing dialog to
     * provide a reason for reporting a message.
     */
    options: string[];
    [property: string]: any;
}

export interface ExtensionRibbonsArrayTabsItem {
    /**
     * A unique identifier for this tab within the app. Maximum length is 64 characters.
     */
    id?: string;
    /**
     * Displayed text for the tab. Maximum length is 64 characters.
     */
    label?:    string;
    position?: Position;
    /**
     * Id of the existing office Tab. Maximum length is 64 characters.
     */
    builtInTabId?: string;
    /**
     * Defines tab groups.
     */
    groups?: ExtensionRibbonsCustomTabGroupsItem[];
    /**
     * Defines mobile group item.
     */
    customMobileRibbonGroups?: ExtensionRibbonsCustomMobileGroupItem[];
}

export interface ExtensionRibbonsCustomMobileGroupItem {
    /**
     * Specify the Id of the group. Used for mobileMessageRead ext point.
     */
    id: string;
    /**
     * Short label of the control. Maximum length is 32 characters.
     */
    label:    string;
    controls: ExtensionRibbonsCustomMobileControlButtonItem[];
    [property: string]: any;
}

export interface ExtensionRibbonsCustomMobileControlButtonItem {
    /**
     * Specify the Id of the button like msgReadFunctionButton.
     */
    id:   string;
    type: "mobileButton";
    /**
     * Short label of the control. Maximum length is 32 characters.
     */
    label: string;
    icons: ExtensionCustomMobileIcon[];
    /**
     * The ID of an action defined in runtimes. Maximum length is 64 characters.
     */
    actionId: string;
    [property: string]: any;
}

export interface ExtensionCustomMobileIcon {
    /**
     * Size in pixels of the icon. Three image sizes are required (25, 32, and 48 pixels).
     */
    size: number;
    /**
     * Url to the icon.
     */
    url: string;
    /**
     * How to scale - 1,2,3 for each image. This attribute specifies the UIScreen.scale property
     * for iOS devices.
     */
    scale: number;
}

export interface ExtensionRibbonsCustomTabGroupsItem {
    /**
     * A unique identifier for this group within the app. Maximum length is 64 characters.
     */
    id?: string;
    /**
     * Displayed text for the group. Maximum length is 64 characters.
     */
    label?:    string;
    icons?:    ExtensionCommonIcon[];
    controls?: ExtensionCommonCustomGroupControlsItem[];
    /**
     * Id of a built-in Group. Maximum length is 64 characters.
     */
    builtInGroupId?: string;
    /**
     * Specifies whether a group will be hidden on application and platform combinations that
     * support the API (Office.ribbon.requestCreateControls) that installs custom contextual
     * tabs on the ribbon. Default is false.
     */
    overriddenByRibbonApi?: boolean;
}

export interface Position {
    /**
     * The id of the built-in tab. Maximum length is 64 characters.
     */
    builtInTabId: string;
    /**
     * Define alignment of this custom tab relative to the specified built-in tab.
     */
    align: Align;
}

/**
 * Define alignment of this custom tab relative to the specified built-in tab.
 */
export type Align = "after" | "before";

/**
 * General runtime for "MailApp" or "TaskpaneApp". Configures the set of runtimes and
 * actions that can be used by each extension point. Min size 1.
 *
 * A runtime environment for a page or script.
 */
export interface ExtensionRuntimesArray {
    requirements?: RequirementsExtensionElement;
    /**
     * A unique identifier for this runtime within the app.  Maximum length is 64 characters.
     */
    id: string;
    /**
     * Supports running functions and launching pages.
     */
    type?: "general";
    code:  ExtensionRuntimeCode;
    /**
     * Runtimes with a short lifetime do not preserve state across executions. Runtimes with a
     * long lifetime do.
     */
    lifetime?:        Lifetime;
    actions?:         ExtensionRuntimesActionsItem[];
    customFunctions?: ExtensionCustomFunctions;
}

/**
 * Specifies the set of actions supported by this runtime. An action is either running a
 * JavaScript function or opening a view such as a task pane.
 */
export interface ExtensionRuntimesActionsItem {
    /**
     * Identifier for this action. Maximum length is 64 characters. This value is passed to the
     * code file.
     */
    id: string;
    /**
     * executeFunction: Run a script function without waiting for it to finish. openPage: Open a
     * page in a view. executeDataFunction: invoke command and retrieve data.
     */
    type: ActionType;
    /**
     * Display name of the action. Maximum length is 64 characters.
     */
    displayName?: string;
    /**
     * Specifies that a task pane supports pinning, which keeps the task pane open when the user
     * changes the selection.
     */
    pinnable?: boolean;
    /**
     * View where the page should be opened. Maximum length is 64 characters.
     */
    view?: string;
    /**
     * Whether allows the action to have multiple selection.
     */
    multiselect?: boolean;
    /**
     * Whether allows task pane add-ins to activate without the Reading Pane enabled or a
     * message selected.
     */
    supportsNoItemContext?: boolean;
}

/**
 * executeFunction: Run a script function without waiting for it to finish. openPage: Open a
 * page in a view. executeDataFunction: invoke command and retrieve data.
 */
export type ActionType = "executeFunction" | "openPage" | "executeDataFunction";

/**
 * Custom function enable developers to add new functions to Excel by defining those
 * functions in JavaScript as part of an add-in. Users within Excel can access custom
 * functions just as they would any native function in Excel, such as SUM().
 */
export interface ExtensionCustomFunctions {
    /**
     * Array of function object which defines function metadata.
     */
    functions: ExtensionFunction[];
    namespace: ExtensionCustomFunctionsNamespace;
    /**
     * Allows a custom function to accept Excel data types as parameters and return values.
     */
    allowCustomDataForDataTypeAny?: boolean;
    [property: string]: any;
}

export interface ExtensionFunction {
    /**
     * A unique ID for the function.
     */
    id: string;
    /**
     * The name of the function that end users see in Excel. In Excel, this function name is
     * prefixed by the custom functions namespace that's specified in the manifest file.
     */
    name: string;
    /**
     * The description of the function that end users see in Excel.
     */
    description?: string;
    /**
     * URL that provides information about the function. (It is displayed in a task pane.)
     */
    helpUrl?: string;
    /**
     * Array that defines the input parameters for the function.
     */
    parameters: ExtensionFunctionParameter[];
    result:     ExtensionResult;
    /**
     * If true, the function can output repeatedly to the cell even when invoked only once. This
     * option is useful for rapidly-changing data sources, such as a stock price. The function
     * should have no return statement. Instead, the result value is passed as the argument of
     * the StreamingInvocation.setResult callback function.
     */
    stream?: boolean;
    /**
     * If true, the function recalculates each time Excel recalculates, instead of only when the
     * formula's dependent values have changed. A function can't use both the stream and
     * volatile properties. If the stream and volatile properties are both set to true, the
     * volatile property will be ignored.
     */
    volatile?: boolean;
    /**
     * If true, Excel calls the CancelableInvocation handler whenever the user takes an action
     * that has the effect of canceling the function; for example, manually triggering
     * recalculation or editing a cell that is referenced by the function. Cancelable functions
     * are typically only used for asynchronous functions that return a single result and need
     * to handle the cancellation of a request for data. A function can't use both the stream
     * and cancelable properties.
     */
    cancelable?: boolean;
    /**
     * If true, your custom function can access the address of the cell that invoked it. The
     * address property of the invocation parameter contains the address of the cell that
     * invoked your custom function. A function can't use both the stream and requiresAddress
     * properties.
     */
    requiresAddress?: boolean;
    /**
     * If true, your custom function can access the addresses of the function's input
     * parameters. This property must be used in combination with the dimensionality property of
     * the result object, and dimensionality must be set to matrix.
     */
    requiresParameterAddress?: boolean;
    [property: string]: any;
}

export interface ExtensionFunctionParameter {
    /**
     * The name of the parameter. This name is displayed in Excel's IntelliSense.
     */
    name: string;
    /**
     * A description of the parameter. This is displayed in Excel's IntelliSense.
     */
    description?: string;
    /**
     * The data type of the parameter. It can only be boolean, number, string, any,
     * CustomFunctions.Invocation, CustomFunctions.StreamingInvocation or
     * CustomFunctions.CancelableInvocation, any allows you to use any of other types.
     */
    type?: string;
    /**
     * A subfield of the type property. Specifies the Excel data types accepted by the custom
     * function. Accepts the values cellvalue, booleancellvalue, doublecellvalue,
     * entitycellvalue, errorcellvalue, formattednumbercellvalue, linkedentitycellvalue,
     * localimagecellvalue, stringcellvalue, webimagecellvalue
     */
    cellValueType?: CellValueType;
    /**
     * Must be either scalar (a non-array value) or matrix (a 2-dimensional array).
     */
    dimensionality?: Dimensionality;
    /**
     * If true, the parameter is optional.
     */
    optional?: boolean;
    /**
     * If true, parameters populate from a specified array. Note that functions all repeating
     * parameters are considered optional parameters by definition.
     */
    repeating?: boolean;
    [property: string]: any;
}

/**
 * A subfield of the type property. Specifies the Excel data types accepted by the custom
 * function. Accepts the values cellvalue, booleancellvalue, doublecellvalue,
 * entitycellvalue, errorcellvalue, formattednumbercellvalue, linkedentitycellvalue,
 * localimagecellvalue, stringcellvalue, webimagecellvalue
 */
export type CellValueType = "cellvalue" | "booleancellvalue" | "doublecellvalue" | "entitycellvalue" | "errorcellvalue" | "formattednumbercellvalue" | "linkedentitycellvalue" | "localimagecellvalue" | "stringcellvalue" | "webimagecellvalue";

/**
 * Must be either scalar (a non-array value) or matrix (a 2-dimensional array).
 *
 * Must be either scalar (a non-array value) or matrix (a 2-dimensional array). Default:
 * scalar.
 */
export type Dimensionality = "scalar" | "matrix";

/**
 * Object that defines the type of information that is returned by the function.
 */
export interface ExtensionResult {
    /**
     * Must be either scalar (a non-array value) or matrix (a 2-dimensional array). Default:
     * scalar.
     */
    dimensionality?: Dimensionality;
    [property: string]: any;
}

/**
 * Defines the namespace for your custom functions. A namespace prepends itself to your
 * custom functions to help customers identify your functions as part of your add-in.
 */
export interface ExtensionCustomFunctionsNamespace {
    /**
     * Non-localizeable version of the namespace.
     */
    id: string;
    /**
     * Localizeable version of the namespace.
     */
    name: string;
    [property: string]: any;
}

/**
 * Runtimes with a short lifetime do not preserve state across executions. Runtimes with a
 * long lifetime do.
 */
export type Lifetime = "short" | "long";

/**
 * Supports running functions and launching pages.
 */

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
     * A relative file path to a transparent PNG outline icon. The border color needs to be
     * white. Size 32x32.
     */
    outline: string;
    /**
     * A relative file path to a full color PNG icon. Size 192x192.
     */
    color: string;
    /**
     * A relative file path to a full color PNG icon with transparent background. Size 32x32.
     */
    color32x32?: string;
}

/**
 * The Intune-related properties for the app.
 */
export interface IntuneInfo {
    /**
     * Supported mobile app managment version that the app is compliant with.
     */
    supportedMobileAppManagementVersion?: string;
}

export interface LocalizationInfo {
    /**
     * The language tag of the strings in this top level manifest file.
     */
    defaultLanguageTag: string;
    /**
     * A relative file path to a the .json file containing strings in the default language.
     */
    defaultLanguageFile?: string;
    additionalLanguages?: AdditionalLanguage[];
}

export interface AdditionalLanguage {
    /**
     * The language tag of the strings in the provided file.
     */
    languageTag: string;
    /**
     * A relative file path to a the .json file containing the translated strings.
     */
    file: string;
}

/**
 * The version of the schema this manifest is using.
 */
export type ManifestVersion = "devPreview" | "m365DevPreview";

/**
 * Specify meeting extension definition.
 */
export interface MeetingExtensionDefinition {
    /**
     * Meeting supported scenes.
     */
    scenes?: Scene[];
    /**
     * Represents if the app has added support for sharing to stage.
     */
    supportsCustomShareToStage?: boolean;
    /**
     * Meeting supported video filters.
     */
    videoFilters?: VideoFilter[];
    /**
     * A URL for configuring the video filters.
     */
    videoFiltersConfigurationUrl?: string;
    /**
     * A boolean value indicating whether this app can stream the meeting's audio video content
     * to an RTMP endpoint.
     */
    supportsStreaming?: boolean;
    /**
     * A boolean value indicating whether this app supports access by anonymous guest users.
     */
    supportsAnonymousGuestUsers?: boolean;
}

export interface Scene {
    /**
     * A unique identifier for this scene. This id must be a GUID.
     */
    id: string;
    /**
     * Scene name.
     */
    name: string;
    /**
     * A relative file path to a scene metadata json file.
     */
    file: string;
    /**
     * A relative file path to a scene PNG preview icon.
     */
    preview: string;
    /**
     * Maximum audiences supported in scene.
     */
    maxAudience: number;
    /**
     * Number of seats reserved for organizers or presenters.
     */
    seatsReservedForOrganizersOrPresenters: number;
}

export interface VideoFilter {
    /**
     * A unique identifier for this video filter. This id must be a GUID.
     */
    id: string;
    /**
     * Video filter's name.
     */
    name: string;
    /**
     * A relative file path to a video filter's thumbnail.
     */
    thumbnail: string;
}

export interface NameClass {
    /**
     * A short display name for the app.
     */
    short: string;
    /**
     * The full name of the app, used if the full app name exceeds 30 characters.
     */
    full: string;
    /**
     * An abbreviated name for the app.
     */
    abbreviated?: string;
}

export type Permission = "identity" | "messageTeamMembers";

export interface ScopeConstraints {
    /**
     * A list of team thread ids to which your app is restricted to
     */
    teams?: TeamElement[];
    /**
     * A list of chat thread ids to which your app is restricted to
     */
    groupChats?: GroupChatElement[];
}

export interface GroupChatElement {
    /**
     * Chat's thread Id
     */
    id: string;
}

export interface TeamElement {
    /**
     * Team's thread Id
     */
    id: string;
}

export interface StaticTab {
    /**
     * A unique identifier for the entity which the tab displays.
     */
    entityId: string;
    /**
     * The display name of the tab.
     */
    name?: string;
    /**
     * The url which points to the entity UI to be displayed in the Teams canvas.
     */
    contentUrl?: string;
    /**
     * The Microsoft App ID specified for the bot in the Bot Framework portal
     * (https://dev.botframework.com/bots)
     */
    contentBotId?: string;
    /**
     * The url to point at if a user opts to view in a browser.
     */
    websiteUrl?: string;
    /**
     * The url to direct a user's search queries.
     */
    searchUrl?: string;
    /**
     * Specifies whether the tab offers an experience in the context of a channel in a team, or
     * an experience scoped to an individual user alone or a group chat. These options are
     * non-exclusive. Currently static tabs are only supported in the 'personal' scope.
     */
    scopes: StaticTabScope[];
    /**
     * The set of contextItem scopes that a tab belong to
     */
    context?: StaticTabContext[];
    /**
     * The set of supportedPlatform scopes that a tab belong to
     */
    supportedPlatform?: SupportedPlatform[];
    /**
     * The set of requirements for the tab.
     */
    requirementSet?: ElementRequirementSet;
}

export type StaticTabContext = "personalTab" | "channelTab" | "privateChatTab" | "meetingChatTab" | "meetingDetailsTab" | "meetingSidePanel" | "meetingStage" | "teamLevelApp";

export type StaticTabScope = "team" | "personal" | "groupChat";

/**
 * Subscription offer associated with this app.
 */
export interface SubscriptionOffer {
    /**
     * A unique identifier for the Commercial Marketplace Software as a Service Offer.
     */
    offerId: string;
}

export type SupportedChannelType = "sharedChannels" | "privateChannels";

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
    public static toTeamsManifestVDevPreview(json: string): TeamsManifestVDevPreview {
        return cast(JSON.parse(json), r("TeamsManifestVDevPreview"));
    }

    public static teamsManifestVDevPreviewToJson(value: TeamsManifestVDevPreview): string {
        return JSON.stringify(uncast(value, r("TeamsManifestVDevPreview")), null, 4);
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
    "TeamsManifestVDevPreview": o([
        { json: "$schema", js: "$schema", typ: u(undefined, "") },
        { json: "manifestVersion", js: "manifestVersion", typ: r("ManifestVersion") },
        { json: "version", js: "version", typ: "" },
        { json: "id", js: "id", typ: "" },
        { json: "packageName", js: "packageName", typ: u(undefined, "") },
        { json: "localizationInfo", js: "localizationInfo", typ: u(undefined, r("LocalizationInfo")) },
        { json: "developer", js: "developer", typ: r("Developer") },
        { json: "name", js: "name", typ: r("NameClass") },
        { json: "description", js: "description", typ: r("Description") },
        { json: "icons", js: "icons", typ: r("Icons") },
        { json: "accentColor", js: "accentColor", typ: "" },
        { json: "configurableTabs", js: "configurableTabs", typ: u(undefined, a(r("ConfigurableTab"))) },
        { json: "staticTabs", js: "staticTabs", typ: u(undefined, a(r("StaticTab"))) },
        { json: "bots", js: "bots", typ: u(undefined, a(r("Bot"))) },
        { json: "connectors", js: "connectors", typ: u(undefined, a(r("Connector"))) },
        { json: "subscriptionOffer", js: "subscriptionOffer", typ: u(undefined, r("SubscriptionOffer")) },
        { json: "composeExtensions", js: "composeExtensions", typ: u(undefined, a(r("ComposeExtension"))) },
        { json: "scopeConstraints", js: "scopeConstraints", typ: u(undefined, r("ScopeConstraints")) },
        { json: "permissions", js: "permissions", typ: u(undefined, a(r("Permission"))) },
        { json: "devicePermissions", js: "devicePermissions", typ: u(undefined, a(r("DevicePermission"))) },
        { json: "validDomains", js: "validDomains", typ: u(undefined, a("")) },
        { json: "webApplicationInfo", js: "webApplicationInfo", typ: u(undefined, r("WebApplicationInfo")) },
        { json: "graphConnector", js: "graphConnector", typ: u(undefined, r("GraphConnector")) },
        { json: "showLoadingIndicator", js: "showLoadingIndicator", typ: u(undefined, true) },
        { json: "isFullScreen", js: "isFullScreen", typ: u(undefined, true) },
        { json: "activities", js: "activities", typ: u(undefined, r("Activities")) },
        { json: "supportedChannelTypes", js: "supportedChannelTypes", typ: u(undefined, a(r("SupportedChannelType"))) },
        { json: "configurableProperties", js: "configurableProperties", typ: u(undefined, a(r("ConfigurableProperty"))) },
        { json: "defaultBlockUntilAdminAction", js: "defaultBlockUntilAdminAction", typ: u(undefined, true) },
        { json: "publisherDocsUrl", js: "publisherDocsUrl", typ: u(undefined, "") },
        { json: "defaultInstallScope", js: "defaultInstallScope", typ: u(undefined, r("DefaultInstallScope")) },
        { json: "defaultGroupCapability", js: "defaultGroupCapability", typ: u(undefined, r("DefaultGroupCapability")) },
        { json: "meetingExtensionDefinition", js: "meetingExtensionDefinition", typ: u(undefined, r("MeetingExtensionDefinition")) },
        { json: "authorization", js: "authorization", typ: u(undefined, r("TeamsManifestVDevPreviewAuthorization")) },
        { json: "extensions", js: "extensions", typ: u(undefined, a(r("ElementExtension"))) },
        { json: "actions", js: "actions", typ: u(undefined, a(r("ElementAction"))) },
        { json: "dashboardCards", js: "dashboardCards", typ: u(undefined, a(r("DashboardCard"))) },
        { json: "intuneInfo", js: "intuneInfo", typ: u(undefined, r("IntuneInfo")) },
        { json: "copilotAgents", js: "copilotAgents", typ: u(undefined, r("CopilotAgents")) },
        { json: "elementRelationshipSet", js: "elementRelationshipSet", typ: u(undefined, r("ElementRelationshipSet")) },
        { json: "backgroundLoadConfiguration", js: "backgroundLoadConfiguration", typ: u(undefined, r("BackgroundLoadConfiguration")) },
    ], false),
    "ElementAction": o([
        { json: "id", js: "id", typ: "" },
        { json: "intent", js: "intent", typ: r("Intent") },
        { json: "displayName", js: "displayName", typ: "" },
        { json: "description", js: "description", typ: "" },
        { json: "icons", js: "icons", typ: u(undefined, a(r("Icon"))) },
        { json: "handlers", js: "handlers", typ: a(r("Handler")) },
    ], false),
    "Handler": o([
        { json: "type", js: "type", typ: r("HandlerType") },
        { json: "supportedObjects", js: "supportedObjects", typ: u(undefined, r("SupportedObjects")) },
        { json: "supportsMultiSelect", js: "supportsMultiSelect", typ: u(undefined, true) },
        { json: "pageInfo", js: "pageInfo", typ: u(undefined, r("PageInfo")) },
        { json: "dialogInfo", js: "dialogInfo", typ: u(undefined, r("DialogInfo")) },
        { json: "url", js: "url", typ: u(undefined, "") },
        { json: "botInfo", js: "botInfo", typ: u(undefined, r("BotInfo")) },
    ], "any"),
    "BotInfo": o([
        { json: "botId", js: "botId", typ: "" },
        { json: "fetchTask", js: "fetchTask", typ: u(undefined, true) },
    ], "any"),
    "DialogInfo": o([
        { json: "dialogType", js: "dialogType", typ: r("DialogType") },
        { json: "url", js: "url", typ: u(undefined, "") },
        { json: "width", js: "width", typ: "" },
        { json: "height", js: "height", typ: "" },
        { json: "parameters", js: "parameters", typ: u(undefined, a(r("ParameterObject"))) },
        { json: "title", js: "title", typ: u(undefined, "") },
    ], "any"),
    "ParameterObject": o([
        { json: "name", js: "name", typ: "" },
        { json: "title", js: "title", typ: "" },
        { json: "description", js: "description", typ: "" },
        { json: "inputType", js: "inputType", typ: "" },
    ], "any"),
    "PageInfo": o([
        { json: "pageId", js: "pageId", typ: "" },
        { json: "subpageId", js: "subpageId", typ: u(undefined, "") },
    ], "any"),
    "SupportedObjects": o([
        { json: "file", js: "file", typ: u(undefined, r("File")) },
        { json: "folder", js: "folder", typ: u(undefined, u(m("any"), null)) },
    ], "any"),
    "File": o([
        { json: "extensions", js: "extensions", typ: u(undefined, a("")) },
    ], false),
    "Icon": o([
        { json: "size", js: "size", typ: 3.14 },
        { json: "url", js: "url", typ: "" },
    ], false),
    "Activities": o([
        { json: "activityTypes", js: "activityTypes", typ: u(undefined, a(r("ActivityType"))) },
        { json: "activityIcons", js: "activityIcons", typ: u(undefined, a(r("ActivityIcon"))) },
    ], false),
    "ActivityIcon": o([
        { json: "id", js: "id", typ: "" },
        { json: "iconFile", js: "iconFile", typ: "" },
    ], false),
    "ActivityType": o([
        { json: "type", js: "type", typ: "" },
        { json: "description", js: "description", typ: "" },
        { json: "templateText", js: "templateText", typ: "" },
        { json: "allowedIconIds", js: "allowedIconIds", typ: u(undefined, a("")) },
    ], false),
    "TeamsManifestVDevPreviewAuthorization": o([
        { json: "permissions", js: "permissions", typ: u(undefined, r("Permissions")) },
    ], false),
    "Permissions": o([
        { json: "resourceSpecific", js: "resourceSpecific", typ: u(undefined, a(r("ResourceSpecific"))) },
    ], false),
    "ResourceSpecific": o([
        { json: "name", js: "name", typ: "" },
        { json: "type", js: "type", typ: r("ResourceSpecificType") },
    ], false),
    "BackgroundLoadConfiguration": o([
        { json: "tabConfiguration", js: "tabConfiguration", typ: u(undefined, r("TabConfiguration")) },
    ], false),
    "TabConfiguration": o([
        { json: "contentUrl", js: "contentUrl", typ: "" },
    ], false),
    "Bot": o([
        { json: "botId", js: "botId", typ: "" },
        { json: "configuration", js: "configuration", typ: u(undefined, r("Configuration")) },
        { json: "allowBotMessageDeleteByUser", js: "allowBotMessageDeleteByUser", typ: u(undefined, true) },
        { json: "needsChannelSelector", js: "needsChannelSelector", typ: u(undefined, true) },
        { json: "isNotificationOnly", js: "isNotificationOnly", typ: u(undefined, true) },
        { json: "requiresSecurityEnabledGroup", js: "requiresSecurityEnabledGroup", typ: u(undefined, true) },
        { json: "supportsFiles", js: "supportsFiles", typ: u(undefined, true) },
        { json: "supportsCalling", js: "supportsCalling", typ: u(undefined, true) },
        { json: "supportsVideo", js: "supportsVideo", typ: u(undefined, true) },
        { json: "scopes", js: "scopes", typ: a(r("CommandListScope")) },
        { json: "commandLists", js: "commandLists", typ: u(undefined, a(r("CommandList"))) },
        { json: "requirementSet", js: "requirementSet", typ: u(undefined, r("ElementRequirementSet")) },
    ], false),
    "CommandList": o([
        { json: "scopes", js: "scopes", typ: a(r("CommandListScope")) },
        { json: "commands", js: "commands", typ: a(r("CommandListCommand")) },
    ], false),
    "CommandListCommand": o([
        { json: "title", js: "title", typ: "" },
        { json: "description", js: "description", typ: "" },
    ], false),
    "Configuration": o([
        { json: "team", js: "team", typ: u(undefined, r("ConfigurationTeam")) },
        { json: "groupChat", js: "groupChat", typ: u(undefined, r("ConfigurationGroupChat")) },
    ], false),
    "ConfigurationGroupChat": o([
        { json: "fetchTask", js: "fetchTask", typ: u(undefined, true) },
        { json: "taskInfo", js: "taskInfo", typ: u(undefined, r("TaskInfo")) },
    ], false),
    "TaskInfo": o([
        { json: "title", js: "title", typ: u(undefined, "") },
        { json: "width", js: "width", typ: u(undefined, "") },
        { json: "height", js: "height", typ: u(undefined, "") },
        { json: "url", js: "url", typ: u(undefined, "") },
    ], false),
    "ConfigurationTeam": o([
        { json: "fetchTask", js: "fetchTask", typ: u(undefined, true) },
        { json: "taskInfo", js: "taskInfo", typ: u(undefined, r("TaskInfo")) },
    ], false),
    "ElementRequirementSet": o([
        { json: "hostMustSupportFunctionalities", js: "hostMustSupportFunctionalities", typ: a(r("HostFunctionality")) },
    ], false),
    "HostFunctionality": o([
        { json: "name", js: "name", typ: r("HostMustSupportFunctionalityName") },
    ], false),
    "ComposeExtension": o([
        { json: "id", js: "id", typ: u(undefined, "") },
        { json: "botId", js: "botId", typ: u(undefined, "") },
        { json: "composeExtensionType", js: "composeExtensionType", typ: u(undefined, r("ComposeExtensionType")) },
        { json: "authorization", js: "authorization", typ: u(undefined, r("ComposeExtensionAuthorization")) },
        { json: "apiSpecificationFile", js: "apiSpecificationFile", typ: u(undefined, "") },
        { json: "canUpdateConfiguration", js: "canUpdateConfiguration", typ: u(undefined, u(true, null)) },
        { json: "commands", js: "commands", typ: u(undefined, a(r("ComposeExtensionCommand"))) },
        { json: "messageHandlers", js: "messageHandlers", typ: u(undefined, a(r("MessageHandler"))) },
        { json: "requirementSet", js: "requirementSet", typ: u(undefined, r("ElementRequirementSet")) },
    ], false),
    "ComposeExtensionAuthorization": o([
        { json: "authType", js: "authType", typ: u(undefined, r("AuthType")) },
        { json: "microsoftEntraConfiguration", js: "microsoftEntraConfiguration", typ: u(undefined, r("MicrosoftEntraConfiguration")) },
        { json: "apiSecretServiceAuthConfiguration", js: "apiSecretServiceAuthConfiguration", typ: u(undefined, r("APISecretServiceAuthConfiguration")) },
        { json: "oAuthConfiguration", js: "oAuthConfiguration", typ: u(undefined, r("OAuthConfiguration")) },
    ], false),
    "APISecretServiceAuthConfiguration": o([
        { json: "apiSecretRegistrationId", js: "apiSecretRegistrationId", typ: u(undefined, "") },
    ], false),
    "MicrosoftEntraConfiguration": o([
        { json: "supportsSingleSignOn", js: "supportsSingleSignOn", typ: u(undefined, true) },
    ], false),
    "OAuthConfiguration": o([
        { json: "oAuthConfigurationId", js: "oAuthConfigurationId", typ: u(undefined, "") },
    ], false),
    "ComposeExtensionCommand": o([
        { json: "id", js: "id", typ: "" },
        { json: "type", js: "type", typ: u(undefined, r("CommandType")) },
        { json: "samplePrompts", js: "samplePrompts", typ: u(undefined, a(r("SamplePrompt"))) },
        { json: "apiResponseRenderingTemplateFile", js: "apiResponseRenderingTemplateFile", typ: u(undefined, "") },
        { json: "context", js: "context", typ: u(undefined, a(r("CommandContext"))) },
        { json: "title", js: "title", typ: "" },
        { json: "description", js: "description", typ: u(undefined, "") },
        { json: "initialRun", js: "initialRun", typ: u(undefined, true) },
        { json: "fetchTask", js: "fetchTask", typ: u(undefined, true) },
        { json: "parameters", js: "parameters", typ: u(undefined, a(r("ParameterClass"))) },
        { json: "taskInfo", js: "taskInfo", typ: u(undefined, r("TaskInfo")) },
        { json: "semanticDescription", js: "semanticDescription", typ: u(undefined, "") },
    ], false),
    "ParameterClass": o([
        { json: "name", js: "name", typ: "" },
        { json: "inputType", js: "inputType", typ: u(undefined, r("InputType")) },
        { json: "isRequired", js: "isRequired", typ: u(undefined, true) },
        { json: "title", js: "title", typ: "" },
        { json: "description", js: "description", typ: u(undefined, "") },
        { json: "value", js: "value", typ: u(undefined, "") },
        { json: "choices", js: "choices", typ: u(undefined, a(r("Choice"))) },
        { json: "semanticDescription", js: "semanticDescription", typ: u(undefined, "") },
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
        { json: "supportsAnonymousAccess", js: "supportsAnonymousAccess", typ: u(undefined, true) },
        { json: "supportsAnonymizedPayloads", js: "supportsAnonymizedPayloads", typ: u(undefined, true) },
    ], "any"),
    "ConfigurableTab": o([
        { json: "id", js: "id", typ: u(undefined, "") },
        { json: "configurationUrl", js: "configurationUrl", typ: "" },
        { json: "canUpdateConfiguration", js: "canUpdateConfiguration", typ: u(undefined, true) },
        { json: "scopes", js: "scopes", typ: a(r("ConfigurableTabScope")) },
        { json: "meetingSurfaces", js: "meetingSurfaces", typ: u(undefined, a(r("MeetingSurface"))) },
        { json: "context", js: "context", typ: u(undefined, a(r("ConfigurableTabContext"))) },
        { json: "supportedPlatform", js: "supportedPlatform", typ: u(undefined, a(r("SupportedPlatform"))) },
        { json: "sharePointPreviewImage", js: "sharePointPreviewImage", typ: u(undefined, "") },
        { json: "supportedSharePointHosts", js: "supportedSharePointHosts", typ: u(undefined, a(r("SupportedSharePointHost"))) },
    ], false),
    "Connector": o([
        { json: "connectorId", js: "connectorId", typ: "" },
        { json: "configurationUrl", js: "configurationUrl", typ: u(undefined, "") },
        { json: "scopes", js: "scopes", typ: a(r("ConnectorScope")) },
    ], false),
    "CopilotAgents": o([
        { json: "declarativeAgents", js: "declarativeAgents", typ: u(undefined, a(r("DeclarativeAgentRef"))) },
        { json: "customEngineAgents", js: "customEngineAgents", typ: u(undefined, a(r("CustomEngineAgent"))) },
    ], false),
    "CustomEngineAgent": o([
        { json: "id", js: "id", typ: "" },
        { json: "type", js: "type", typ: r("SourceTypeEnum") },
    ], false),
    "DeclarativeAgentRef": o([
        { json: "id", js: "id", typ: "" },
        { json: "file", js: "file", typ: "" },
    ], false),
    "DashboardCard": o([
        { json: "id", js: "id", typ: "" },
        { json: "displayName", js: "displayName", typ: "" },
        { json: "description", js: "description", typ: "" },
        { json: "pickerGroupId", js: "pickerGroupId", typ: "" },
        { json: "icon", js: "icon", typ: u(undefined, r("DashboardCardIcon")) },
        { json: "contentSource", js: "contentSource", typ: r("DashboardCardContentSource") },
        { json: "defaultSize", js: "defaultSize", typ: r("DefaultSize") },
    ], false),
    "DashboardCardContentSource": o([
        { json: "sourceType", js: "sourceType", typ: u(undefined, r("SourceTypeEnum")) },
        { json: "botConfiguration", js: "botConfiguration", typ: u(undefined, r("BotConfiguration")) },
    ], false),
    "BotConfiguration": o([
        { json: "botId", js: "botId", typ: u(undefined, "") },
    ], false),
    "DashboardCardIcon": o([
        { json: "iconUrl", js: "iconUrl", typ: u(undefined, "") },
        { json: "officeUIFabricIconName", js: "officeUIFabricIconName", typ: u(undefined, "") },
    ], false),
    "DefaultGroupCapability": o([
        { json: "team", js: "team", typ: u(undefined, r("Groupchat")) },
        { json: "groupchat", js: "groupchat", typ: u(undefined, r("Groupchat")) },
        { json: "meetings", js: "meetings", typ: u(undefined, r("Groupchat")) },
    ], false),
    "Description": o([
        { json: "short", js: "short", typ: "" },
        { json: "full", js: "full", typ: "" },
    ], false),
    "Developer": o([
        { json: "name", js: "name", typ: "" },
        { json: "mpnId", js: "mpnId", typ: u(undefined, "") },
        { json: "websiteUrl", js: "websiteUrl", typ: "" },
        { json: "privacyUrl", js: "privacyUrl", typ: "" },
        { json: "termsOfUseUrl", js: "termsOfUseUrl", typ: "" },
        { json: "contactInfo", js: "contactInfo", typ: u(undefined, r("ContactInfo")) },
    ], false),
    "ContactInfo": o([
        { json: "defaultSupport", js: "defaultSupport", typ: r("DefaultSupport") },
    ], "any"),
    "DefaultSupport": o([
        { json: "userEmailsForChatSupport", js: "userEmailsForChatSupport", typ: a("") },
        { json: "emailsForEmailSupport", js: "emailsForEmailSupport", typ: a("") },
    ], "any"),
    "ElementRelationshipSet": o([
        { json: "oneWayDependencies", js: "oneWayDependencies", typ: u(undefined, a(r("OneWayDependency"))) },
        { json: "mutualDependencies", js: "mutualDependencies", typ: u(undefined, a(a(r("ElementReference")))) },
    ], false),
    "ElementReference": o([
        { json: "name", js: "name", typ: r("MutualDependencyName") },
        { json: "id", js: "id", typ: "" },
        { json: "commandIds", js: "commandIds", typ: u(undefined, a("")) },
    ], false),
    "OneWayDependency": o([
        { json: "element", js: "element", typ: r("ElementReference") },
        { json: "dependsOn", js: "dependsOn", typ: a(r("ElementReference")) },
    ], false),
    "ElementExtension": o([
        { json: "requirements", js: "requirements", typ: u(undefined, r("RequirementsExtensionElement")) },
        { json: "runtimes", js: "runtimes", typ: u(undefined, a(r("ExtensionRuntimesArray"))) },
        { json: "ribbons", js: "ribbons", typ: u(undefined, a(r("ExtensionRibbonsArray"))) },
        { json: "autoRunEvents", js: "autoRunEvents", typ: u(undefined, a(r("ExtensionAutoRunEventsArray"))) },
        { json: "alternates", js: "alternates", typ: u(undefined, a(r("ExtensionAlternateVersionsArray"))) },
        { json: "audienceClaimUrl", js: "audienceClaimUrl", typ: u(undefined, "") },
        { json: "appDeeplinks", js: "appDeeplinks", typ: u(undefined, a(r("ExtensionAppDeeplinksArray"))) },
        { json: "contentRuntimes", js: "contentRuntimes", typ: u(undefined, a(r("ExtensionContentRuntimeArray"))) },
        { json: "getStartedMessages", js: "getStartedMessages", typ: u(undefined, a(r("ExtensionGetStartedMessageArray"))) },
        { json: "contextMenus", js: "contextMenus", typ: u(undefined, a(r("ExtensionContextMenuArray"))) },
        { json: "keyboardShortcuts", js: "keyboardShortcuts", typ: u(undefined, a(r("ExtensionKeyboardShortcut"))) },
    ], false),
    "ExtensionAlternateVersionsArray": o([
        { json: "requirements", js: "requirements", typ: u(undefined, r("RequirementsExtensionElement")) },
        { json: "prefer", js: "prefer", typ: u(undefined, r("Prefer")) },
        { json: "hide", js: "hide", typ: u(undefined, r("Hide")) },
        { json: "alternateIcons", js: "alternateIcons", typ: r("AlternateIcons") },
        { json: "xllCustomFunctions", js: "xllCustomFunctions", typ: u(undefined, r("ExtensionXllCustomFunctions")) },
    ], false),
    "AlternateIcons": o([
        { json: "icon", js: "icon", typ: r("ExtensionCommonIcon") },
        { json: "highResolutionIcon", js: "highResolutionIcon", typ: r("ExtensionCommonIcon") },
    ], false),
    "ExtensionCommonIcon": o([
        { json: "size", js: "size", typ: 3.14 },
        { json: "url", js: "url", typ: "" },
    ], false),
    "Hide": o([
        { json: "storeOfficeAddin", js: "storeOfficeAddin", typ: u(undefined, r("StoreOfficeAddin")) },
        { json: "customOfficeAddin", js: "customOfficeAddin", typ: u(undefined, r("CustomOfficeAddin")) },
    ], "any"),
    "CustomOfficeAddin": o([
        { json: "officeAddinId", js: "officeAddinId", typ: "" },
    ], false),
    "StoreOfficeAddin": o([
        { json: "officeAddinId", js: "officeAddinId", typ: "" },
        { json: "assetId", js: "assetId", typ: "" },
    ], false),
    "Prefer": o([
        { json: "comAddin", js: "comAddin", typ: u(undefined, r("COMAddin")) },
    ], "any"),
    "COMAddin": o([
        { json: "progId", js: "progId", typ: "" },
    ], false),
    "RequirementsExtensionElement": o([
        { json: "capabilities", js: "capabilities", typ: u(undefined, a(r("Capability"))) },
        { json: "scopes", js: "scopes", typ: u(undefined, a(r("RequirementsScope"))) },
        { json: "formFactors", js: "formFactors", typ: u(undefined, a(r("FormFactor"))) },
    ], false),
    "Capability": o([
        { json: "name", js: "name", typ: "" },
        { json: "minVersion", js: "minVersion", typ: u(undefined, "") },
        { json: "maxVersion", js: "maxVersion", typ: u(undefined, "") },
    ], false),
    "ExtensionXllCustomFunctions": o([
        { json: "fileName", js: "fileName", typ: u(undefined, "") },
    ], "any"),
    "ExtensionAppDeeplinksArray": o([
        { json: "requirements", js: "requirements", typ: u(undefined, r("AppDeeplinkRequirements")) },
        { json: "contexts", js: "contexts", typ: a(r("ExtensionContext")) },
        { json: "actionId", js: "actionId", typ: "" },
        { json: "label", js: "label", typ: "" },
        { json: "semanticDescription", js: "semanticDescription", typ: "" },
    ], false),
    "AppDeeplinkRequirements": o([
        { json: "capabilities", js: "capabilities", typ: u(undefined, a(r("Capability"))) },
        { json: "scopes", js: "scopes", typ: u(undefined, a(r("RequirementsScope"))) },
        { json: "formFactors", js: "formFactors", typ: u(undefined, a(r("FormFactor"))) },
    ], false),
    "ExtensionAutoRunEventsArray": o([
        { json: "requirements", js: "requirements", typ: u(undefined, r("RequirementsExtensionElement")) },
        { json: "events", js: "events", typ: a(r("Event")) },
    ], false),
    "Event": o([
        { json: "type", js: "type", typ: "" },
        { json: "actionId", js: "actionId", typ: "" },
        { json: "options", js: "options", typ: u(undefined, r("Options")) },
    ], false),
    "Options": o([
        { json: "sendMode", js: "sendMode", typ: r("SendMode") },
    ], false),
    "ExtensionContentRuntimeArray": o([
        { json: "requirements", js: "requirements", typ: u(undefined, r("ContentRuntimeRequirements")) },
        { json: "id", js: "id", typ: "" },
        { json: "code", js: "code", typ: r("ExtensionRuntimeCode") },
        { json: "requestedHeight", js: "requestedHeight", typ: u(undefined, 3.14) },
        { json: "requestedWidth", js: "requestedWidth", typ: u(undefined, 3.14) },
        { json: "disableSnapshot", js: "disableSnapshot", typ: u(undefined, true) },
    ], false),
    "ExtensionRuntimeCode": o([
        { json: "page", js: "page", typ: "" },
        { json: "script", js: "script", typ: u(undefined, "") },
    ], false),
    "ContentRuntimeRequirements": o([
        { json: "capabilities", js: "capabilities", typ: u(undefined, a(r("Capability"))) },
        { json: "scopes", js: "scopes", typ: u(undefined, a(r("RequirementsScope"))) },
        { json: "formFactors", js: "formFactors", typ: u(undefined, a(r("FormFactor"))) },
    ], false),
    "ExtensionContextMenuArray": o([
        { json: "requirements", js: "requirements", typ: u(undefined, r("ContextMenuRequirements")) },
        { json: "menus", js: "menus", typ: a(r("ExtensionMenuItem")) },
    ], false),
    "ExtensionMenuItem": o([
        { json: "entryPoint", js: "entryPoint", typ: r("EntryPoint") },
        { json: "controls", js: "controls", typ: a(r("ExtensionCommonCustomGroupControlsItem")) },
    ], false),
    "ExtensionCommonCustomGroupControlsItem": o([
        { json: "id", js: "id", typ: "" },
        { json: "type", js: "type", typ: r("PurpleType") },
        { json: "builtInControlId", js: "builtInControlId", typ: u(undefined, "") },
        { json: "label", js: "label", typ: "" },
        { json: "icons", js: "icons", typ: a(r("ExtensionCommonIcon")) },
        { json: "supertip", js: "supertip", typ: r("ExtensionCommonSuperToolTip") },
        { json: "actionId", js: "actionId", typ: u(undefined, "") },
        { json: "overriddenByRibbonApi", js: "overriddenByRibbonApi", typ: u(undefined, true) },
        { json: "enabled", js: "enabled", typ: u(undefined, true) },
        { json: "items", js: "items", typ: u(undefined, a(r("ExtensionCommonCustomControlMenuItem"))) },
    ], false),
    "ExtensionCommonCustomControlMenuItem": o([
        { json: "id", js: "id", typ: "" },
        { json: "type", js: "type", typ: r("ItemType") },
        { json: "label", js: "label", typ: "" },
        { json: "icons", js: "icons", typ: u(undefined, a(r("ExtensionCommonIcon"))) },
        { json: "supertip", js: "supertip", typ: r("ExtensionCommonSuperToolTip") },
        { json: "actionId", js: "actionId", typ: "" },
        { json: "enabled", js: "enabled", typ: u(undefined, true) },
        { json: "overriddenByRibbonApi", js: "overriddenByRibbonApi", typ: u(undefined, true) },
    ], false),
    "ExtensionCommonSuperToolTip": o([
        { json: "title", js: "title", typ: "" },
        { json: "description", js: "description", typ: "" },
    ], false),
    "ContextMenuRequirements": o([
        { json: "capabilities", js: "capabilities", typ: u(undefined, a(r("Capability"))) },
        { json: "scopes", js: "scopes", typ: u(undefined, a(r("RequirementsScope"))) },
        { json: "formFactors", js: "formFactors", typ: u(undefined, a(r("FormFactor"))) },
    ], false),
    "ExtensionGetStartedMessageArray": o([
        { json: "requirements", js: "requirements", typ: u(undefined, r("GetStartedMessageRequirements")) },
        { json: "title", js: "title", typ: "" },
        { json: "description", js: "description", typ: "" },
        { json: "learnMoreUrl", js: "learnMoreUrl", typ: "" },
    ], false),
    "GetStartedMessageRequirements": o([
        { json: "capabilities", js: "capabilities", typ: u(undefined, a(r("Capability"))) },
        { json: "scopes", js: "scopes", typ: u(undefined, a(r("RequirementsScope"))) },
        { json: "formFactors", js: "formFactors", typ: u(undefined, a(r("FormFactor"))) },
    ], false),
    "ExtensionKeyboardShortcut": o([
        { json: "requirements", js: "requirements", typ: u(undefined, r("RequirementsExtensionElement")) },
        { json: "shortcuts", js: "shortcuts", typ: a(r("ExtensionShortcut")) },
    ], "any"),
    "ExtensionShortcut": o([
        { json: "key", js: "key", typ: r("Key") },
        { json: "actionId", js: "actionId", typ: "" },
    ], "any"),
    "Key": o([
        { json: "default", js: "default", typ: "" },
        { json: "mac", js: "mac", typ: u(undefined, "") },
        { json: "web", js: "web", typ: u(undefined, "") },
        { json: "windows", js: "windows", typ: u(undefined, "") },
    ], "any"),
    "ExtensionRibbonsArray": o([
        { json: "requirements", js: "requirements", typ: u(undefined, r("RequirementsExtensionElement")) },
        { json: "contexts", js: "contexts", typ: u(undefined, a(r("ExtensionContext"))) },
        { json: "tabs", js: "tabs", typ: a(r("ExtensionRibbonsArrayTabsItem")) },
        { json: "fixedControls", js: "fixedControls", typ: u(undefined, a(r("ExtensionRibbonsArrayFixedControlItem"))) },
        { json: "spamPreProcessingDialog", js: "spamPreProcessingDialog", typ: u(undefined, r("ExtensionRibbonsSpamPreProcessingDialog")) },
    ], false),
    "ExtensionRibbonsArrayFixedControlItem": o([
        { json: "id", js: "id", typ: "" },
        { json: "type", js: "type", typ: r("FixedControlType") },
        { json: "label", js: "label", typ: "" },
        { json: "icons", js: "icons", typ: a(r("ExtensionCommonIcon")) },
        { json: "supertip", js: "supertip", typ: r("ExtensionCommonSuperToolTip") },
        { json: "actionId", js: "actionId", typ: "" },
        { json: "enabled", js: "enabled", typ: true },
    ], false),
    "ExtensionRibbonsSpamPreProcessingDialog": o([
        { json: "title", js: "title", typ: "" },
        { json: "description", js: "description", typ: "" },
        { json: "spamReportingOptions", js: "spamReportingOptions", typ: u(undefined, r("SpamReportingOptions")) },
        { json: "spamFreeTextSectionTitle", js: "spamFreeTextSectionTitle", typ: u(undefined, "") },
        { json: "spamMoreInfo", js: "spamMoreInfo", typ: u(undefined, r("SpamMoreInfo")) },
    ], false),
    "SpamMoreInfo": o([
        { json: "text", js: "text", typ: "" },
        { json: "url", js: "url", typ: "" },
    ], "any"),
    "SpamReportingOptions": o([
        { json: "title", js: "title", typ: "" },
        { json: "options", js: "options", typ: a("") },
    ], "any"),
    "ExtensionRibbonsArrayTabsItem": o([
        { json: "id", js: "id", typ: u(undefined, "") },
        { json: "label", js: "label", typ: u(undefined, "") },
        { json: "position", js: "position", typ: u(undefined, r("Position")) },
        { json: "builtInTabId", js: "builtInTabId", typ: u(undefined, "") },
        { json: "groups", js: "groups", typ: u(undefined, a(r("ExtensionRibbonsCustomTabGroupsItem"))) },
        { json: "customMobileRibbonGroups", js: "customMobileRibbonGroups", typ: u(undefined, a(r("ExtensionRibbonsCustomMobileGroupItem"))) },
    ], false),
    "ExtensionRibbonsCustomMobileGroupItem": o([
        { json: "id", js: "id", typ: "" },
        { json: "label", js: "label", typ: "" },
        { json: "controls", js: "controls", typ: a(r("ExtensionRibbonsCustomMobileControlButtonItem")) },
    ], "any"),
    "ExtensionRibbonsCustomMobileControlButtonItem": o([
        { json: "id", js: "id", typ: "" },
        { json: "type", js: "type", typ: r("FluffyType") },
        { json: "label", js: "label", typ: "" },
        { json: "icons", js: "icons", typ: a(r("ExtensionCustomMobileIcon")) },
        { json: "actionId", js: "actionId", typ: "" },
    ], "any"),
    "ExtensionCustomMobileIcon": o([
        { json: "size", js: "size", typ: 3.14 },
        { json: "url", js: "url", typ: "" },
        { json: "scale", js: "scale", typ: 3.14 },
    ], false),
    "ExtensionRibbonsCustomTabGroupsItem": o([
        { json: "id", js: "id", typ: u(undefined, "") },
        { json: "label", js: "label", typ: u(undefined, "") },
        { json: "icons", js: "icons", typ: u(undefined, a(r("ExtensionCommonIcon"))) },
        { json: "controls", js: "controls", typ: u(undefined, a(r("ExtensionCommonCustomGroupControlsItem"))) },
        { json: "builtInGroupId", js: "builtInGroupId", typ: u(undefined, "") },
        { json: "overriddenByRibbonApi", js: "overriddenByRibbonApi", typ: u(undefined, true) },
    ], false),
    "Position": o([
        { json: "builtInTabId", js: "builtInTabId", typ: "" },
        { json: "align", js: "align", typ: r("Align") },
    ], false),
    "ExtensionRuntimesArray": o([
        { json: "requirements", js: "requirements", typ: u(undefined, r("RequirementsExtensionElement")) },
        { json: "id", js: "id", typ: "" },
        { json: "type", js: "type", typ: u(undefined, r("RuntimeType")) },
        { json: "code", js: "code", typ: r("ExtensionRuntimeCode") },
        { json: "lifetime", js: "lifetime", typ: u(undefined, r("Lifetime")) },
        { json: "actions", js: "actions", typ: u(undefined, a(r("ExtensionRuntimesActionsItem"))) },
        { json: "customFunctions", js: "customFunctions", typ: u(undefined, r("ExtensionCustomFunctions")) },
    ], false),
    "ExtensionRuntimesActionsItem": o([
        { json: "id", js: "id", typ: "" },
        { json: "type", js: "type", typ: r("ActionType") },
        { json: "displayName", js: "displayName", typ: u(undefined, "") },
        { json: "pinnable", js: "pinnable", typ: u(undefined, true) },
        { json: "view", js: "view", typ: u(undefined, "") },
        { json: "multiselect", js: "multiselect", typ: u(undefined, true) },
        { json: "supportsNoItemContext", js: "supportsNoItemContext", typ: u(undefined, true) },
    ], false),
    "ExtensionCustomFunctions": o([
        { json: "functions", js: "functions", typ: a(r("ExtensionFunction")) },
        { json: "namespace", js: "namespace", typ: r("ExtensionCustomFunctionsNamespace") },
        { json: "allowCustomDataForDataTypeAny", js: "allowCustomDataForDataTypeAny", typ: u(undefined, true) },
    ], "any"),
    "ExtensionFunction": o([
        { json: "id", js: "id", typ: "" },
        { json: "name", js: "name", typ: "" },
        { json: "description", js: "description", typ: u(undefined, "") },
        { json: "helpUrl", js: "helpUrl", typ: u(undefined, "") },
        { json: "parameters", js: "parameters", typ: a(r("ExtensionFunctionParameter")) },
        { json: "result", js: "result", typ: r("ExtensionResult") },
        { json: "stream", js: "stream", typ: u(undefined, true) },
        { json: "volatile", js: "volatile", typ: u(undefined, true) },
        { json: "cancelable", js: "cancelable", typ: u(undefined, true) },
        { json: "requiresAddress", js: "requiresAddress", typ: u(undefined, true) },
        { json: "requiresParameterAddress", js: "requiresParameterAddress", typ: u(undefined, true) },
    ], "any"),
    "ExtensionFunctionParameter": o([
        { json: "name", js: "name", typ: "" },
        { json: "description", js: "description", typ: u(undefined, "") },
        { json: "type", js: "type", typ: u(undefined, "") },
        { json: "cellValueType", js: "cellValueType", typ: u(undefined, r("CellValueType")) },
        { json: "dimensionality", js: "dimensionality", typ: u(undefined, r("Dimensionality")) },
        { json: "optional", js: "optional", typ: u(undefined, true) },
        { json: "repeating", js: "repeating", typ: u(undefined, true) },
    ], "any"),
    "ExtensionResult": o([
        { json: "dimensionality", js: "dimensionality", typ: u(undefined, r("Dimensionality")) },
    ], "any"),
    "ExtensionCustomFunctionsNamespace": o([
        { json: "id", js: "id", typ: "" },
        { json: "name", js: "name", typ: "" },
    ], "any"),
    "GraphConnector": o([
        { json: "notificationUrl", js: "notificationUrl", typ: "" },
    ], false),
    "Icons": o([
        { json: "outline", js: "outline", typ: "" },
        { json: "color", js: "color", typ: "" },
        { json: "color32x32", js: "color32x32", typ: u(undefined, "") },
    ], false),
    "IntuneInfo": o([
        { json: "supportedMobileAppManagementVersion", js: "supportedMobileAppManagementVersion", typ: u(undefined, "") },
    ], false),
    "LocalizationInfo": o([
        { json: "defaultLanguageTag", js: "defaultLanguageTag", typ: "" },
        { json: "defaultLanguageFile", js: "defaultLanguageFile", typ: u(undefined, "") },
        { json: "additionalLanguages", js: "additionalLanguages", typ: u(undefined, a(r("AdditionalLanguage"))) },
    ], false),
    "AdditionalLanguage": o([
        { json: "languageTag", js: "languageTag", typ: "" },
        { json: "file", js: "file", typ: "" },
    ], false),
    "MeetingExtensionDefinition": o([
        { json: "scenes", js: "scenes", typ: u(undefined, a(r("Scene"))) },
        { json: "supportsCustomShareToStage", js: "supportsCustomShareToStage", typ: u(undefined, true) },
        { json: "videoFilters", js: "videoFilters", typ: u(undefined, a(r("VideoFilter"))) },
        { json: "videoFiltersConfigurationUrl", js: "videoFiltersConfigurationUrl", typ: u(undefined, "") },
        { json: "supportsStreaming", js: "supportsStreaming", typ: u(undefined, true) },
        { json: "supportsAnonymousGuestUsers", js: "supportsAnonymousGuestUsers", typ: u(undefined, true) },
    ], false),
    "Scene": o([
        { json: "id", js: "id", typ: "" },
        { json: "name", js: "name", typ: "" },
        { json: "file", js: "file", typ: "" },
        { json: "preview", js: "preview", typ: "" },
        { json: "maxAudience", js: "maxAudience", typ: 0 },
        { json: "seatsReservedForOrganizersOrPresenters", js: "seatsReservedForOrganizersOrPresenters", typ: 0 },
    ], false),
    "VideoFilter": o([
        { json: "id", js: "id", typ: "" },
        { json: "name", js: "name", typ: "" },
        { json: "thumbnail", js: "thumbnail", typ: "" },
    ], false),
    "NameClass": o([
        { json: "short", js: "short", typ: "" },
        { json: "full", js: "full", typ: "" },
        { json: "abbreviated", js: "abbreviated", typ: u(undefined, "") },
    ], false),
    "ScopeConstraints": o([
        { json: "teams", js: "teams", typ: u(undefined, a(r("TeamElement"))) },
        { json: "groupChats", js: "groupChats", typ: u(undefined, a(r("GroupChatElement"))) },
    ], false),
    "GroupChatElement": o([
        { json: "id", js: "id", typ: "" },
    ], false),
    "TeamElement": o([
        { json: "id", js: "id", typ: "" },
    ], false),
    "StaticTab": o([
        { json: "entityId", js: "entityId", typ: "" },
        { json: "name", js: "name", typ: u(undefined, "") },
        { json: "contentUrl", js: "contentUrl", typ: u(undefined, "") },
        { json: "contentBotId", js: "contentBotId", typ: u(undefined, "") },
        { json: "websiteUrl", js: "websiteUrl", typ: u(undefined, "") },
        { json: "searchUrl", js: "searchUrl", typ: u(undefined, "") },
        { json: "scopes", js: "scopes", typ: a(r("StaticTabScope")) },
        { json: "context", js: "context", typ: u(undefined, a(r("StaticTabContext"))) },
        { json: "supportedPlatform", js: "supportedPlatform", typ: u(undefined, a(r("SupportedPlatform"))) },
        { json: "requirementSet", js: "requirementSet", typ: u(undefined, r("ElementRequirementSet")) },
    ], false),
    "SubscriptionOffer": o([
        { json: "offerId", js: "offerId", typ: "" },
    ], false),
    "WebApplicationInfo": o([
        { json: "id", js: "id", typ: "" },
        { json: "resource", js: "resource", typ: u(undefined, "") },
    ], false),
    "DialogType": [
        "adaptiveCard",
        "url",
    ],
    "HandlerType": [
        "invokeAPI",
        "invokeBot",
        "openDialog",
        "openPage",
        "openTaskpane",
        "openURL",
    ],
    "Intent": [
        "addTo",
        "create",
        "custom",
        "open",
        "preview",
        "share",
        "sign",
    ],
    "ResourceSpecificType": [
        "Application",
        "Delegated",
    ],
    "CommandListScope": [
        "copilot",
        "groupChat",
        "personal",
        "team",
    ],
    "HostMustSupportFunctionalityName": [
        "dialogAdaptiveCard",
        "dialogAdaptiveCardBot",
        "dialogUrl",
        "dialogUrlBot",
    ],
    "AuthType": [
        "apiSecretServiceAuth",
        "microsoftEntra",
        "none",
        "oAuth2.0",
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
        "callingSidePanel",
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
    "SupportedPlatform": [
        "desktop",
        "mobile",
        "teamsMeetingDevices",
    ],
    "SupportedSharePointHost": [
        "sharePointFullPage",
        "sharePointWebPart",
    ],
    "ConnectorScope": [
        "team",
    ],
    "SourceTypeEnum": [
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
        "copilot",
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
    "MutualDependencyName": [
        "bots",
        "composeExtensions",
        "configurableTabs",
        "staticTabs",
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
    "ExtensionContext": [
        "default",
        "logEventMeetingDetailsAttendee",
        "mailCompose",
        "mailRead",
        "meetingDetailsAttendee",
        "meetingDetailsOrganizer",
        "onlineMeetingDetailsOrganizer",
        "spamReportingOverride",
    ],
    "SendMode": [
        "block",
        "promptUser",
        "softBlock",
    ],
    "ItemType": [
        "menuItem",
    ],
    "PurpleType": [
        "button",
        "menu",
    ],
    "EntryPoint": [
        "cell",
        "text",
    ],
    "FixedControlType": [
        "button",
    ],
    "FluffyType": [
        "mobileButton",
    ],
    "Align": [
        "after",
        "before",
    ],
    "ActionType": [
        "executeDataFunction",
        "executeFunction",
        "openPage",
    ],
    "CellValueType": [
        "booleancellvalue",
        "cellvalue",
        "doublecellvalue",
        "entitycellvalue",
        "errorcellvalue",
        "formattednumbercellvalue",
        "linkedentitycellvalue",
        "localimagecellvalue",
        "stringcellvalue",
        "webimagecellvalue",
    ],
    "Dimensionality": [
        "matrix",
        "scalar",
    ],
    "Lifetime": [
        "long",
        "short",
    ],
    "RuntimeType": [
        "general",
    ],
    "ManifestVersion": [
        "devPreview",
        "m365DevPreview",
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
    "StaticTabScope": [
        "groupChat",
        "personal",
        "team",
    ],
    "SupportedChannelType": [
        "privateChannels",
        "sharedChannels",
    ],
};
