// To parse this data:
//
//   import { Convert, TeamsManifestV1D6 } from "./file";
//
//   const teamsManifestV1D6 = Convert.toTeamsManifestV1D6(json);
//
// These functions will throw an error if the JSON doesn't
// match the expected interface, even if the JSON is valid.

export interface TeamsManifestV1D6 {
    $schema?: string;
    /**
     * The version of the schema this manifest is using.
     */
    manifestVersion: "1.6";
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
    name:              Name;
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
     * The set of compose extensions for this app. Currently only one compose extension per app
     * is supported.
     */
    composeExtensions?: ComposeExtension[];
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
     * A value indicating whether or not show loading indicator when app/tab is loading
     */
    showLoadingIndicator?: boolean;
}

export interface Bot {
    /**
     * The Microsoft App ID specified for the bot in the Bot Framework portal
     * (https://dev.botframework.com/bots)
     */
    botId: string;
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
     * A value indicating whether the bot supports uploading/downloading of files.
     */
    supportsFiles?: boolean;
    /**
     * Specifies whether the bot offers an experience in the context of a channel in a team, in
     * a 1:1 or group chat, or in an experience scoped to an individual user alone. These
     * options are non-exclusive.
     */
    scopes: CommandListScope[];
    /**
     * The list of commands that the bot supplies, including their usage, description, and the
     * scope for which the commands are valid. A separate command list should be used for each
     * scope.
     */
    commandLists?: CommandList[];
}

export interface CommandList {
    /**
     * Specifies the scopes for which the command list is valid
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

export type CommandListScope = "team" | "personal" | "groupChat" | "groupchat";

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
     * Id of the command.
     */
    id: string;
    /**
     * Type of the command
     */
    type?: CommandType;
    /**
     * Context where the command would apply
     */
    context?: Context[];
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
    parameters?: Parameter[];
    taskInfo?:   TaskInfo;
}

export type Context = "compose" | "commandBox" | "message";

export interface Parameter {
    /**
     * Name of the parameter.
     */
    name: string;
    /**
     * Type of the parameter
     */
    inputType?: InputType;
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

export interface TaskInfo {
    /**
     * Initial dialog title
     */
    title?: string;
    /**
     * Dialog width - either a number in pixels or default layout such as 'large', 'medium', or
     * 'small'
     */
    width?: string;
    /**
     * Dialog height - either a number in pixels or default layout such as 'large', 'medium', or
     * 'small'
     */
    height?: string;
    /**
     * Initial webview URL
     */
    url?: string;
}

/**
 * Type of the command
 */
export type CommandType = "query" | "action";

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
    [property: string]: any;
}

export interface ConfigurableTab {
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
     * A relative file path to a tab preview image for use in SharePoint. Size 1024x768.
     */
    sharePointPreviewImage?: string;
    /**
     * Defines how your tab will be made available in SharePoint.
     */
    supportedSharePointHosts?: SupportedSharePointHost[];
}

export type ConfigurableTabScope = "team" | "groupChat" | "groupchat";

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
}

export type DevicePermission = "geolocation" | "media" | "notifications" | "midi" | "openExternal";

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
}

export interface LocalizationInfo {
    /**
     * The language tag of the strings in this top level manifest file.
     */
    defaultLanguageTag:   string;
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

export interface Name {
    /**
     * A short display name for the app.
     */
    short: string;
    /**
     * The full name of the app, used if the full app name exceeds 30 characters.
     */
    full?: string;
}

export type Permission = "identity" | "messageTeamMembers";

export interface StaticTab {
    /**
     * A unique identifier for the entity which the tab displays.
     */
    entityId: string;
    /**
     * The display name of the tab.
     */
    name: string;
    /**
     * The url which points to the entity UI to be displayed in the Teams canvas.
     */
    contentUrl: string;
    /**
     * The url to point at if a user opts to view in a browser.
     */
    websiteUrl?: string;
    /**
     * Specifies whether the tab offers an experience in the context of a channel in a team, or
     * an experience scoped to an individual user alone. These options are non-exclusive.
     * Currently static tabs are only supported in the 'personal' scope.
     */
    scopes: StaticTabScope[];
}

export type StaticTabScope = "team" | "personal";

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
    resource?:               string;
    applicationPermissions?: string[];
}

// Converts JSON strings to/from your types
// and asserts the results of JSON.parse at runtime
export class Convert {
    public static toTeamsManifestV1D6(json: string): TeamsManifestV1D6 {
        return cast(JSON.parse(json), r("TeamsManifestV1D6"));
    }

    public static teamsManifestV1D6ToJson(value: TeamsManifestV1D6): string {
        return JSON.stringify(uncast(value, r("TeamsManifestV1D6")), null, 4);
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
    "TeamsManifestV1D6": o([
        { json: "$schema", js: "$schema", typ: u(undefined, "") },
        { json: "manifestVersion", js: "manifestVersion", typ: r("ManifestVersion") },
        { json: "version", js: "version", typ: "" },
        { json: "id", js: "id", typ: "" },
        { json: "packageName", js: "packageName", typ: u(undefined, "") },
        { json: "localizationInfo", js: "localizationInfo", typ: u(undefined, r("LocalizationInfo")) },
        { json: "developer", js: "developer", typ: r("Developer") },
        { json: "name", js: "name", typ: r("Name") },
        { json: "description", js: "description", typ: r("Description") },
        { json: "icons", js: "icons", typ: r("Icons") },
        { json: "accentColor", js: "accentColor", typ: "" },
        { json: "configurableTabs", js: "configurableTabs", typ: u(undefined, a(r("ConfigurableTab"))) },
        { json: "staticTabs", js: "staticTabs", typ: u(undefined, a(r("StaticTab"))) },
        { json: "bots", js: "bots", typ: u(undefined, a(r("Bot"))) },
        { json: "connectors", js: "connectors", typ: u(undefined, a(r("Connector"))) },
        { json: "composeExtensions", js: "composeExtensions", typ: u(undefined, a(r("ComposeExtension"))) },
        { json: "permissions", js: "permissions", typ: u(undefined, a(r("Permission"))) },
        { json: "devicePermissions", js: "devicePermissions", typ: u(undefined, a(r("DevicePermission"))) },
        { json: "validDomains", js: "validDomains", typ: u(undefined, a("")) },
        { json: "webApplicationInfo", js: "webApplicationInfo", typ: u(undefined, r("WebApplicationInfo")) },
        { json: "showLoadingIndicator", js: "showLoadingIndicator", typ: u(undefined, true) },
    ], false),
    "Bot": o([
        { json: "botId", js: "botId", typ: "" },
        { json: "needsChannelSelector", js: "needsChannelSelector", typ: u(undefined, true) },
        { json: "isNotificationOnly", js: "isNotificationOnly", typ: u(undefined, true) },
        { json: "supportsFiles", js: "supportsFiles", typ: u(undefined, true) },
        { json: "scopes", js: "scopes", typ: a(r("CommandListScope")) },
        { json: "commandLists", js: "commandLists", typ: u(undefined, a(r("CommandList"))) },
    ], false),
    "CommandList": o([
        { json: "scopes", js: "scopes", typ: a(r("CommandListScope")) },
        { json: "commands", js: "commands", typ: a(r("CommandListCommand")) },
    ], false),
    "CommandListCommand": o([
        { json: "title", js: "title", typ: "" },
        { json: "description", js: "description", typ: "" },
    ], false),
    "ComposeExtension": o([
        { json: "botId", js: "botId", typ: "" },
        { json: "canUpdateConfiguration", js: "canUpdateConfiguration", typ: u(undefined, true) },
        { json: "commands", js: "commands", typ: a(r("ComposeExtensionCommand")) },
        { json: "messageHandlers", js: "messageHandlers", typ: u(undefined, a(r("MessageHandler"))) },
    ], false),
    "ComposeExtensionCommand": o([
        { json: "id", js: "id", typ: "" },
        { json: "type", js: "type", typ: u(undefined, r("CommandType")) },
        { json: "context", js: "context", typ: u(undefined, a(r("Context"))) },
        { json: "title", js: "title", typ: "" },
        { json: "description", js: "description", typ: u(undefined, "") },
        { json: "initialRun", js: "initialRun", typ: u(undefined, true) },
        { json: "fetchTask", js: "fetchTask", typ: u(undefined, true) },
        { json: "parameters", js: "parameters", typ: u(undefined, a(r("Parameter"))) },
        { json: "taskInfo", js: "taskInfo", typ: u(undefined, r("TaskInfo")) },
    ], false),
    "Parameter": o([
        { json: "name", js: "name", typ: "" },
        { json: "inputType", js: "inputType", typ: u(undefined, r("InputType")) },
        { json: "title", js: "title", typ: "" },
        { json: "description", js: "description", typ: u(undefined, "") },
        { json: "value", js: "value", typ: u(undefined, "") },
        { json: "choices", js: "choices", typ: u(undefined, a(r("Choice"))) },
    ], false),
    "Choice": o([
        { json: "title", js: "title", typ: "" },
        { json: "value", js: "value", typ: "" },
    ], false),
    "TaskInfo": o([
        { json: "title", js: "title", typ: u(undefined, "") },
        { json: "width", js: "width", typ: u(undefined, "") },
        { json: "height", js: "height", typ: u(undefined, "") },
        { json: "url", js: "url", typ: u(undefined, "") },
    ], false),
    "MessageHandler": o([
        { json: "type", js: "type", typ: r("MessageHandlerType") },
        { json: "value", js: "value", typ: r("Value") },
    ], false),
    "Value": o([
        { json: "domains", js: "domains", typ: u(undefined, a("")) },
    ], "any"),
    "ConfigurableTab": o([
        { json: "configurationUrl", js: "configurationUrl", typ: "" },
        { json: "canUpdateConfiguration", js: "canUpdateConfiguration", typ: u(undefined, true) },
        { json: "scopes", js: "scopes", typ: a(r("ConfigurableTabScope")) },
        { json: "sharePointPreviewImage", js: "sharePointPreviewImage", typ: u(undefined, "") },
        { json: "supportedSharePointHosts", js: "supportedSharePointHosts", typ: u(undefined, a(r("SupportedSharePointHost"))) },
    ], false),
    "Connector": o([
        { json: "connectorId", js: "connectorId", typ: "" },
        { json: "configurationUrl", js: "configurationUrl", typ: u(undefined, "") },
        { json: "scopes", js: "scopes", typ: a(r("ConnectorScope")) },
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
    ], false),
    "Icons": o([
        { json: "outline", js: "outline", typ: "" },
        { json: "color", js: "color", typ: "" },
    ], false),
    "LocalizationInfo": o([
        { json: "defaultLanguageTag", js: "defaultLanguageTag", typ: "" },
        { json: "additionalLanguages", js: "additionalLanguages", typ: u(undefined, a(r("AdditionalLanguage"))) },
    ], false),
    "AdditionalLanguage": o([
        { json: "languageTag", js: "languageTag", typ: "" },
        { json: "file", js: "file", typ: "" },
    ], false),
    "Name": o([
        { json: "short", js: "short", typ: "" },
        { json: "full", js: "full", typ: u(undefined, "") },
    ], false),
    "StaticTab": o([
        { json: "entityId", js: "entityId", typ: "" },
        { json: "name", js: "name", typ: "" },
        { json: "contentUrl", js: "contentUrl", typ: "" },
        { json: "websiteUrl", js: "websiteUrl", typ: u(undefined, "") },
        { json: "scopes", js: "scopes", typ: a(r("StaticTabScope")) },
    ], false),
    "WebApplicationInfo": o([
        { json: "id", js: "id", typ: "" },
        { json: "resource", js: "resource", typ: u(undefined, "") },
        { json: "applicationPermissions", js: "applicationPermissions", typ: u(undefined, a("")) },
    ], false),
    "CommandListScope": [
        "groupChat",
        "groupchat",
        "personal",
        "team",
    ],
    "Context": [
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
    "DevicePermission": [
        "geolocation",
        "midi",
        "media",
        "notifications",
        "openExternal",
    ],
    "ManifestVersion": [
        "1.6",
    ],
    "Permission": [
        "identity",
        "messageTeamMembers",
    ],
    "StaticTabScope": [
        "personal",
        "team",
    ],
};
