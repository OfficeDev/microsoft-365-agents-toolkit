// To parse this data:
//
//   import { Convert, MicrosoftTeamsV1D0 } from "./file";
//
//   const microsoftTeamsV1D0 = Convert.toMicrosoftTeamsV1D0(json);
//
// These functions will throw an error if the JSON doesn't
// match the expected interface, even if the JSON is valid.

export interface MicrosoftTeamsV1D0 {
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
     * These are tabs users can optionally add to their channels and require extra configuration
     * before they are added. Configurable tabs are not supported in the personal scope.
     * Currently only one configurable tab per app is supported.
     */
    configurableTabs?: ConfigurableTab[];
    /**
     * The set of Office365 connectors for this app. Currently only one connector per app is
     * supported.
     */
    connectors?: Connector[];
    description: Description;
    developer:   Developer;
    icons:       Icons;
    /**
     * A unique identifier for this app. This id must be a GUID.
     */
    id: string;
    /**
     * The version of the schema this manifest is using.
     */
    manifestVersion: string;
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
}

export interface Bot {
    /**
     * The Microsoft App ID specified for the bot in the Bot Framework portal
     * (https://dev.botframework.com/bots)
     */
    botId: string;
    /**
     * The list of commands that the bot supplies, including their usage, description, and the
     * scope for which the commands are valid. A seperate command list should be used for each
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
     * Specifies whether the bot offers an experience in the context of a channel in a team, or
     * an experience scoped to an individual user alone. These options are non-exclusive.
     */
    scopes: CommandListScope[];
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

/**
 * Specifies whether the compose extension offers an experience in the context of a channel
 * in a team or an experience scoped to an individual user alone. These options are
 * non-exclusive.
 */
export enum CommandListScope {
    Personal = "personal",
    Team = "team",
}

export interface ComposeExtension {
    /**
     * The Microsoft App ID specified for the bot powering the compose extension in the Bot
     * Framework portal (https://dev.botframework.com/bots)
     */
    botId:    string;
    commands: ComposeExtensionCommand[];
    scopes:   CommandListScope[];
}

export interface ComposeExtensionCommand {
    /**
     * Description of the command.
     */
    description?: string;
    /**
     * Id of the command.
     */
    id: string;
    /**
     * A boolean value that indicates if the command should be run once initially with no
     * parameter.
     */
    initialRun?: boolean;
    parameters:  Parameter[];
    /**
     * Title of the command.
     */
    title: string;
}

export interface Parameter {
    /**
     * Description of the parameter.
     */
    description?: string;
    /**
     * Name of the parameter.
     */
    name: string;
    /**
     * Title of the parameter.
     */
    title: string;
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
     * Specifies whether the tab offers an experience in the context of a channel in a team, or
     * an experience scoped to an individual user alone. These options are non-exclusive.
     * Currently, configurable tabs are only supported in the teams scope.
     */
    scopes: ConfigurableTabScope[];
}

export enum ConfigurableTabScope {
    Team = "team",
}

export interface Connector {
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
    scopes: ConfigurableTabScope[];
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
    contentUrl: string;
    /**
     * A unique identifier for the entity which the tab displays.
     */
    entityId: string;
    /**
     * The display name of the tab.
     */
    name: string;
    /**
     * Specifies whether the tab offers an experience in the context of a channel in a team, or
     * an experience scoped to an individual user alone. These options are non-exclusive.
     * Currently static tabs are only supported in the 'personal' scope.
     */
    scopes: CommandListScope[];
    /**
     * The url to point at if a user opts to view in a browser.
     */
    websiteUrl?: string;
}

// Converts JSON strings to/from your types
// and asserts the results of JSON.parse at runtime
export class Convert {
    public static toMicrosoftTeamsV1D0(json: string): MicrosoftTeamsV1D0 {
        return cast(JSON.parse(json), r("MicrosoftTeamsV1D0"));
    }

    public static microsoftTeamsV1D0ToJson(value: MicrosoftTeamsV1D0): string {
        return JSON.stringify(uncast(value, r("MicrosoftTeamsV1D0")), null, 2);
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
    "MicrosoftTeamsV1D0": o([
        { json: "accentColor", js: "accentColor", typ: "" },
        { json: "bots", js: "bots", typ: u(undefined, a(r("Bot"))) },
        { json: "composeExtensions", js: "composeExtensions", typ: u(undefined, a(r("ComposeExtension"))) },
        { json: "configurableTabs", js: "configurableTabs", typ: u(undefined, a(r("ConfigurableTab"))) },
        { json: "connectors", js: "connectors", typ: u(undefined, a(r("Connector"))) },
        { json: "description", js: "description", typ: r("Description") },
        { json: "developer", js: "developer", typ: r("Developer") },
        { json: "icons", js: "icons", typ: r("Icons") },
        { json: "id", js: "id", typ: "" },
        { json: "manifestVersion", js: "manifestVersion", typ: "" },
        { json: "name", js: "name", typ: r("Name") },
        { json: "packageName", js: "packageName", typ: "" },
        { json: "permissions", js: "permissions", typ: u(undefined, a(r("Permission"))) },
        { json: "staticTabs", js: "staticTabs", typ: u(undefined, a(r("StaticTab"))) },
        { json: "validDomains", js: "validDomains", typ: u(undefined, a("")) },
        { json: "version", js: "version", typ: "" },
    ], false),
    "Bot": o([
        { json: "botId", js: "botId", typ: "" },
        { json: "commandLists", js: "commandLists", typ: u(undefined, a(r("CommandList"))) },
        { json: "isNotificationOnly", js: "isNotificationOnly", typ: u(undefined, true) },
        { json: "needsChannelSelector", js: "needsChannelSelector", typ: u(undefined, true) },
        { json: "scopes", js: "scopes", typ: a(r("CommandListScope")) },
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
        { json: "commands", js: "commands", typ: a(r("ComposeExtensionCommand")) },
        { json: "scopes", js: "scopes", typ: a(r("CommandListScope")) },
    ], false),
    "ComposeExtensionCommand": o([
        { json: "description", js: "description", typ: u(undefined, "") },
        { json: "id", js: "id", typ: "" },
        { json: "initialRun", js: "initialRun", typ: u(undefined, true) },
        { json: "parameters", js: "parameters", typ: a(r("Parameter")) },
        { json: "title", js: "title", typ: "" },
    ], false),
    "Parameter": o([
        { json: "description", js: "description", typ: u(undefined, "") },
        { json: "name", js: "name", typ: "" },
        { json: "title", js: "title", typ: "" },
    ], false),
    "ConfigurableTab": o([
        { json: "canUpdateConfiguration", js: "canUpdateConfiguration", typ: u(undefined, true) },
        { json: "configurationUrl", js: "configurationUrl", typ: "" },
        { json: "scopes", js: "scopes", typ: a(r("ConfigurableTabScope")) },
    ], false),
    "Connector": o([
        { json: "connectorId", js: "connectorId", typ: "" },
        { json: "scopes", js: "scopes", typ: a(r("ConfigurableTabScope")) },
    ], false),
    "Description": o([
        { json: "full", js: "full", typ: "" },
        { json: "short", js: "short", typ: "" },
    ], false),
    "Developer": o([
        { json: "name", js: "name", typ: "" },
        { json: "privacyUrl", js: "privacyUrl", typ: "" },
        { json: "termsOfUseUrl", js: "termsOfUseUrl", typ: "" },
        { json: "websiteUrl", js: "websiteUrl", typ: "" },
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
        { json: "contentUrl", js: "contentUrl", typ: "" },
        { json: "entityId", js: "entityId", typ: "" },
        { json: "name", js: "name", typ: "" },
        { json: "scopes", js: "scopes", typ: a(r("CommandListScope")) },
        { json: "websiteUrl", js: "websiteUrl", typ: u(undefined, "") },
    ], false),
    "CommandListScope": [
        "personal",
        "team",
    ],
    "ConfigurableTabScope": [
        "team",
    ],
    "Permission": [
        "identity",
        "messageTeamMembers",
    ],
};
