// To parse this data:
//
//   import { Convert, MicrosoftTeamsV1D14 } from "./file";
//
//   const microsoftTeamsV1D14 = Convert.toMicrosoftTeamsV1D14(json);
//
// These functions will throw an error if the JSON doesn't
// match the expected interface, even if the JSON is valid.

export interface MicrosoftTeamsV1D14 {
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
    authorization?: Authorization;
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
    connectors?: Connector[];
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
     * A unique identifier for this app in reverse domain notation. E.g: com.example.myapp
     */
    packageName?: string;
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
export interface Authorization {
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
    commandLists?: CommandList[];
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
    Groupchat = "groupchat",
    Personal = "personal",
    Team = "team",
}

export interface ComposeExtension {
    /**
     * The Microsoft App ID specified for the bot powering the compose extension in the Bot
     * Framework portal (https://dev.botframework.com/bots)
     */
    botId: string;
    /**
     * A value indicating whether the configuration of a compose extension can be updated by the
     * user.
     */
    canUpdateConfiguration?: boolean;
    commands:                ComposeExtensionCommand[];
    /**
     * A list of handlers that allow apps to be invoked when certain conditions are met
     */
    messageHandlers?: MessageHandler[];
}

export interface ComposeExtensionCommand {
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
    initialRun?: boolean;
    parameters?: Parameter[];
    taskInfo?:   TaskInfo;
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
     * Name of the parameter.
     */
    name: string;
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

/**
 * Type of the command
 */
export enum CommandType {
    Action = "action",
    Query = "query",
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
    CallingSidePanel = "callingSidePanel",
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
    Groupchat = "groupchat",
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
    Groupchat = "groupchat",
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
    The114 = "1.14",
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
    full?: string;
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
     * an experience scoped to an individual user alone. These options are non-exclusive.
     * Currently static tabs are only supported in the 'personal' scope.
     */
    scopes: StaticTabScope[];
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
    PersonalTab = "personalTab",
}

export enum StaticTabScope {
    Personal = "personal",
    Team = "team",
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
    public static toMicrosoftTeamsV1D14(json: string): MicrosoftTeamsV1D14 {
        return cast(JSON.parse(json), r("MicrosoftTeamsV1D14"));
    }

    public static microsoftTeamsV1D14ToJson(value: MicrosoftTeamsV1D14): string {
        return JSON.stringify(uncast(value, r("MicrosoftTeamsV1D14")), null, 2);
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
    "MicrosoftTeamsV1D14": o([
        { json: "$schema", js: "$schema", typ: u(undefined, "") },
        { json: "accentColor", js: "accentColor", typ: "" },
        { json: "activities", js: "activities", typ: u(undefined, r("Activities")) },
        { json: "authorization", js: "authorization", typ: u(undefined, r("Authorization")) },
        { json: "bots", js: "bots", typ: u(undefined, a(r("Bot"))) },
        { json: "composeExtensions", js: "composeExtensions", typ: u(undefined, a(r("ComposeExtension"))) },
        { json: "configurableProperties", js: "configurableProperties", typ: u(undefined, a(r("ConfigurableProperty"))) },
        { json: "configurableTabs", js: "configurableTabs", typ: u(undefined, a(r("ConfigurableTab"))) },
        { json: "connectors", js: "connectors", typ: u(undefined, a(r("Connector"))) },
        { json: "defaultBlockUntilAdminAction", js: "defaultBlockUntilAdminAction", typ: u(undefined, true) },
        { json: "defaultGroupCapability", js: "defaultGroupCapability", typ: u(undefined, r("DefaultGroupCapability")) },
        { json: "defaultInstallScope", js: "defaultInstallScope", typ: u(undefined, r("DefaultInstallScope")) },
        { json: "description", js: "description", typ: r("Description") },
        { json: "developer", js: "developer", typ: r("Developer") },
        { json: "devicePermissions", js: "devicePermissions", typ: u(undefined, a(r("DevicePermission"))) },
        { json: "graphConnector", js: "graphConnector", typ: u(undefined, r("GraphConnector")) },
        { json: "icons", js: "icons", typ: r("Icons") },
        { json: "id", js: "id", typ: "" },
        { json: "isFullScreen", js: "isFullScreen", typ: u(undefined, true) },
        { json: "localizationInfo", js: "localizationInfo", typ: u(undefined, r("LocalizationInfo")) },
        { json: "manifestVersion", js: "manifestVersion", typ: r("ManifestVersion") },
        { json: "meetingExtensionDefinition", js: "meetingExtensionDefinition", typ: u(undefined, r("MeetingExtensionDefinition")) },
        { json: "name", js: "name", typ: r("Name") },
        { json: "packageName", js: "packageName", typ: u(undefined, "") },
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
    "Authorization": o([
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
    "ComposeExtension": o([
        { json: "botId", js: "botId", typ: "" },
        { json: "canUpdateConfiguration", js: "canUpdateConfiguration", typ: u(undefined, true) },
        { json: "commands", js: "commands", typ: a(r("ComposeExtensionCommand")) },
        { json: "messageHandlers", js: "messageHandlers", typ: u(undefined, a(r("MessageHandler"))) },
    ], false),
    "ComposeExtensionCommand": o([
        { json: "context", js: "context", typ: u(undefined, a(r("CommandContext"))) },
        { json: "description", js: "description", typ: u(undefined, "") },
        { json: "fetchTask", js: "fetchTask", typ: u(undefined, true) },
        { json: "id", js: "id", typ: "" },
        { json: "initialRun", js: "initialRun", typ: u(undefined, true) },
        { json: "parameters", js: "parameters", typ: u(undefined, a(r("Parameter"))) },
        { json: "taskInfo", js: "taskInfo", typ: u(undefined, r("TaskInfo")) },
        { json: "title", js: "title", typ: "" },
        { json: "type", js: "type", typ: u(undefined, r("CommandType")) },
    ], false),
    "Parameter": o([
        { json: "choices", js: "choices", typ: u(undefined, a(r("Choice"))) },
        { json: "description", js: "description", typ: u(undefined, "") },
        { json: "inputType", js: "inputType", typ: u(undefined, r("InputType")) },
        { json: "name", js: "name", typ: "" },
        { json: "title", js: "title", typ: "" },
        { json: "value", js: "value", typ: u(undefined, "") },
    ], false),
    "Choice": o([
        { json: "title", js: "title", typ: "" },
        { json: "value", js: "value", typ: "" },
    ], false),
    "TaskInfo": o([
        { json: "height", js: "height", typ: u(undefined, "") },
        { json: "title", js: "title", typ: u(undefined, "") },
        { json: "url", js: "url", typ: u(undefined, "") },
        { json: "width", js: "width", typ: u(undefined, "") },
    ], false),
    "MessageHandler": o([
        { json: "type", js: "type", typ: r("MessageHandlerType") },
        { json: "value", js: "value", typ: r("Value") },
    ], false),
    "Value": o([
        { json: "domains", js: "domains", typ: u(undefined, a("")) },
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
    "GraphConnector": o([
        { json: "notificationUrl", js: "notificationUrl", typ: "" },
    ], false),
    "Icons": o([
        { json: "color", js: "color", typ: "" },
        { json: "outline", js: "outline", typ: "" },
    ], false),
    "LocalizationInfo": o([
        { json: "additionalLanguages", js: "additionalLanguages", typ: u(undefined, a(r("AdditionalLanguage"))) },
        { json: "defaultLanguageTag", js: "defaultLanguageTag", typ: "" },
    ], false),
    "AdditionalLanguage": o([
        { json: "file", js: "file", typ: "" },
        { json: "languageTag", js: "languageTag", typ: "" },
    ], false),
    "MeetingExtensionDefinition": o([
        { json: "scenes", js: "scenes", typ: u(undefined, a(r("Scene"))) },
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
        { json: "full", js: "full", typ: u(undefined, "") },
        { json: "short", js: "short", typ: "" },
    ], false),
    "StaticTab": o([
        { json: "contentBotId", js: "contentBotId", typ: u(undefined, "") },
        { json: "contentUrl", js: "contentUrl", typ: u(undefined, "") },
        { json: "context", js: "context", typ: u(undefined, a(r("StaticTabContext"))) },
        { json: "entityId", js: "entityId", typ: "" },
        { json: "name", js: "name", typ: u(undefined, "") },
        { json: "scopes", js: "scopes", typ: a(r("StaticTabScope")) },
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
        "groupchat",
        "personal",
        "team",
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
        "groupchat",
        "team",
    ],
    "SupportedSharePointHost": [
        "sharePointFullPage",
        "sharePointWebPart",
    ],
    "ConnectorScope": [
        "team",
    ],
    "Groupchat": [
        "bot",
        "connector",
        "tab",
    ],
    "DefaultInstallScope": [
        "groupChat",
        "groupchat",
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
    "ManifestVersion": [
        "1.14",
    ],
    "Permission": [
        "identity",
        "messageTeamMembers",
    ],
    "StaticTabContext": [
        "channelTab",
        "personalTab",
    ],
    "StaticTabScope": [
        "personal",
        "team",
    ],
    "SupportedChannelType": [
        "privateChannels",
        "sharedChannels",
    ],
};
