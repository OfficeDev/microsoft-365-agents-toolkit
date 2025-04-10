// To parse this data:
//
//   import { Convert, MicrosoftTeamsV1D4 } from "./file";
//
//   const microsoftTeamsV1D4 = Convert.toMicrosoftTeamsV1D4(json);
//
// These functions will throw an error if the JSON doesn't
// match the expected interface, even if the JSON is valid.

export interface MicrosoftTeamsV1D4 {
    schema?: string;
    /**
     * A color to use in conjunction with the icon. The value must be a valid HTML color code
     * starting with '#', for example `#4464ee`.
     */
    accentColor: string;
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
    description: Description;
    developer:   Developer;
    /**
     * Specify the native features on a user's device that your app may request access to.
     */
    devicePermissions?: DevicePermission[];
    icons:              Icons;
    /**
     * A unique identifier for this app. This id must be a GUID.
     */
    id: string;
    /**
     * The version of the schema this manifest is using.
     */
    manifestVersion: ManifestVersion;
    name:            Name;
    /**
     * A unique identifier for this app in reverse domain notation. E.g: com.example.myapp
     */
    packageName: string;
    /**
     * Specifies the permissions the app requests from users.
     */
    permissions?: Permission[];
    /**
     * A set of tabs that may be 'pinned' by default, without the user adding them manually.
     * Static tabs declared in personal scope are always pinned to the app's personal
     * experience. Static tabs do not currently support the 'teams' scope.
     */
    staticTabs?: StaticTab[];
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

export interface Bot {
    /**
     * The Microsoft App ID specified for the bot in the Bot Framework portal
     * (https://dev.botframework.com/bots)
     */
    botID: string;
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
     * A value indicating whether the bot supports uploading/downloading of files.
     */
    supportsFiles?: boolean;
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
    botID: string;
    /**
     * A value indicating whether the configuration of a compose extension can be updated by the
     * user.
     */
    canUpdateConfiguration?: boolean;
    commands:                ComposeExtensionCommand[];
}

export interface ComposeExtensionCommand {
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
    /**
     * Title of the command.
     */
    title: string;
    /**
     * Type of the command
     */
    type?: Type;
}

export interface Parameter {
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
}

/**
 * Type of the parameter
 */
export enum InputType {
    Date = "date",
    Number = "number",
    Text = "text",
    Textarea = "textarea",
    Time = "time",
    Toggle = "toggle",
}

/**
 * Type of the command
 */
export enum Type {
    Action = "action",
    Query = "query",
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
    configurationURL: string;
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
    configurationURL?: string;
    /**
     * A unique identifier for the connector which matches its ID in the Connectors Developer
     * Portal.
     */
    connectorID: string;
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
     * The display name for the developer.
     */
    name: string;
    /**
     * The url to the page that provides privacy information for the app.
     */
    privacyURL: string;
    /**
     * The url to the page that provides the terms of use for the app.
     */
    termsOfUseURL: string;
    /**
     * The url to the page that provides support information for the app.
     */
    websiteURL: string;
}

export enum DevicePermission {
    Geolocation = "geolocation",
    MIDI = "midi",
    Media = "media",
    Notifications = "notifications",
    OpenExternal = "openExternal",
}

export interface Icons {
    /**
     * A relative file path to a full color PNG icon. Size 96x96.
     */
    color: string;
    /**
     * A relative file path to a transparent PNG outline icon. The border color needs to be
     * white. Size 20x20.
     */
    outline: string;
}

export enum ManifestVersion {
    The14 = "1.4",
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
     * The url which points to the entity UI to be displayed in the Teams canvas.
     */
    contentURL: string;
    /**
     * A unique identifier for the entity which the tab displays.
     */
    entityID: string;
    /**
     * The display name of the tab.
     */
    name: string;
    /**
     * Specifies whether the tab offers an experience in the context of a channel in a team, or
     * an experience scoped to an individual user alone. These options are non-exclusive.
     * Currently static tabs are only supported in the 'personal' scope.
     */
    scopes: StaticTabScope[];
    /**
     * The url to point at if a user opts to view in a browser.
     */
    websiteURL?: string;
}

export enum StaticTabScope {
    Personal = "personal",
    Team = "team",
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
    resource: string;
}

// Converts JSON strings to/from your types
// and asserts the results of JSON.parse at runtime
export class Convert {
    public static toMicrosoftTeamsV1D4(json: string): MicrosoftTeamsV1D4 {
        return cast(JSON.parse(json), r("MicrosoftTeamsV1D4"));
    }

    public static microsoftTeamsV1D4ToJson(value: MicrosoftTeamsV1D4): string {
        return JSON.stringify(uncast(value, r("MicrosoftTeamsV1D4")), null, 2);
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
    "MicrosoftTeamsV1D4": o([
        { json: "$schema", js: "schema", typ: u(undefined, "") },
        { json: "accentColor", js: "accentColor", typ: "" },
        { json: "bots", js: "bots", typ: u(undefined, a(r("Bot"))) },
        { json: "composeExtensions", js: "composeExtensions", typ: u(undefined, a(r("ComposeExtension"))) },
        { json: "configurableTabs", js: "configurableTabs", typ: u(undefined, a(r("ConfigurableTab"))) },
        { json: "connectors", js: "connectors", typ: u(undefined, a(r("Connector"))) },
        { json: "description", js: "description", typ: r("Description") },
        { json: "developer", js: "developer", typ: r("Developer") },
        { json: "devicePermissions", js: "devicePermissions", typ: u(undefined, a(r("DevicePermission"))) },
        { json: "icons", js: "icons", typ: r("Icons") },
        { json: "id", js: "id", typ: "" },
        { json: "manifestVersion", js: "manifestVersion", typ: r("ManifestVersion") },
        { json: "name", js: "name", typ: r("Name") },
        { json: "packageName", js: "packageName", typ: "" },
        { json: "permissions", js: "permissions", typ: u(undefined, a(r("Permission"))) },
        { json: "staticTabs", js: "staticTabs", typ: u(undefined, a(r("StaticTab"))) },
        { json: "validDomains", js: "validDomains", typ: u(undefined, a("")) },
        { json: "version", js: "version", typ: "" },
        { json: "webApplicationInfo", js: "webApplicationInfo", typ: u(undefined, r("WebApplicationInfo")) },
    ], false),
    "Bot": o([
        { json: "botId", js: "botID", typ: "" },
        { json: "commandLists", js: "commandLists", typ: u(undefined, a(r("CommandList"))) },
        { json: "isNotificationOnly", js: "isNotificationOnly", typ: u(undefined, true) },
        { json: "needsChannelSelector", js: "needsChannelSelector", typ: u(undefined, true) },
        { json: "scopes", js: "scopes", typ: a(r("CommandListScope")) },
        { json: "supportsFiles", js: "supportsFiles", typ: u(undefined, true) },
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
        { json: "botId", js: "botID", typ: "" },
        { json: "canUpdateConfiguration", js: "canUpdateConfiguration", typ: u(undefined, true) },
        { json: "commands", js: "commands", typ: a(r("ComposeExtensionCommand")) },
    ], false),
    "ComposeExtensionCommand": o([
        { json: "description", js: "description", typ: u(undefined, "") },
        { json: "fetchTask", js: "fetchTask", typ: u(undefined, true) },
        { json: "id", js: "id", typ: "" },
        { json: "initialRun", js: "initialRun", typ: u(undefined, true) },
        { json: "parameters", js: "parameters", typ: u(undefined, a(r("Parameter"))) },
        { json: "title", js: "title", typ: "" },
        { json: "type", js: "type", typ: u(undefined, r("Type")) },
    ], false),
    "Parameter": o([
        { json: "description", js: "description", typ: u(undefined, "") },
        { json: "inputType", js: "inputType", typ: u(undefined, r("InputType")) },
        { json: "name", js: "name", typ: "" },
        { json: "title", js: "title", typ: "" },
    ], false),
    "ConfigurableTab": o([
        { json: "canUpdateConfiguration", js: "canUpdateConfiguration", typ: u(undefined, true) },
        { json: "configurationUrl", js: "configurationURL", typ: "" },
        { json: "scopes", js: "scopes", typ: a(r("ConfigurableTabScope")) },
        { json: "sharePointPreviewImage", js: "sharePointPreviewImage", typ: u(undefined, "") },
        { json: "supportedSharePointHosts", js: "supportedSharePointHosts", typ: u(undefined, a(r("SupportedSharePointHost"))) },
    ], false),
    "Connector": o([
        { json: "configurationUrl", js: "configurationURL", typ: u(undefined, "") },
        { json: "connectorId", js: "connectorID", typ: "" },
        { json: "scopes", js: "scopes", typ: a(r("ConnectorScope")) },
    ], false),
    "Description": o([
        { json: "full", js: "full", typ: "" },
        { json: "short", js: "short", typ: "" },
    ], false),
    "Developer": o([
        { json: "name", js: "name", typ: "" },
        { json: "privacyUrl", js: "privacyURL", typ: "" },
        { json: "termsOfUseUrl", js: "termsOfUseURL", typ: "" },
        { json: "websiteUrl", js: "websiteURL", typ: "" },
    ], false),
    "Icons": o([
        { json: "color", js: "color", typ: "" },
        { json: "outline", js: "outline", typ: "" },
    ], false),
    "Name": o([
        { json: "full", js: "full", typ: u(undefined, "") },
        { json: "short", js: "short", typ: "" },
    ], false),
    "StaticTab": o([
        { json: "contentUrl", js: "contentURL", typ: "" },
        { json: "entityId", js: "entityID", typ: "" },
        { json: "name", js: "name", typ: "" },
        { json: "scopes", js: "scopes", typ: a(r("StaticTabScope")) },
        { json: "websiteUrl", js: "websiteURL", typ: u(undefined, "") },
    ], false),
    "WebApplicationInfo": o([
        { json: "id", js: "id", typ: "" },
        { json: "resource", js: "resource", typ: "" },
    ], false),
    "CommandListScope": [
        "groupChat",
        "groupchat",
        "personal",
        "team",
    ],
    "InputType": [
        "date",
        "number",
        "text",
        "textarea",
        "time",
        "toggle",
    ],
    "Type": [
        "action",
        "query",
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
        "1.4",
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
