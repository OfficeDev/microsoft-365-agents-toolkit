// To parse this data:
//
//   import { Convert, DeclarativeAgentManifestV1D3 } from "./file";
//
//   const declarativeAgentManifestV1D3 = Convert.toDeclarativeAgentManifestV1D3(json);
//
// These functions will throw an error if the JSON doesn't
// match the expected interface, even if the JSON is valid.

/**
 * The root of the declarative agent manifest document is a JSON object that contains
 * members that describe the declarative agent.
 */
export interface DeclarativeAgentManifestV1D3 {
    /**
     * Required. Not localizable. The version of the schema this manifest is using.
     */
    version: "v1.3";
    /**
     * Optional. Not localizable.
     */
    id?: string;
    /**
     * Required. Localizable. The name of the declarative agent. It MUST contain at least one
     * nonwhitespace character and MUST be 100 characters or less.
     */
    name: string;
    /**
     * Required. Localizable. The description of the declarative agent. It MUST contain at least
     * one nonwhitespace character and MUST be 1,000 characters or less.
     */
    description: string;
    /**
     * Optional. Not localizable. The detailed instructions or guidelines on how the declarative
     * agent should behave, its functions, and any behaviors to avoid. It MUST contain at least
     * one nonwhitespace character and MUST be 8,000 characters or less.
     */
    instructions?: string;
    /**
     * Optional. Contains an array of objects that define capabilities of the declarative agent.
     */
    capabilities?: CapabilityElement[];
    /**
     * Optional. A list of examples of questions that the declarative agent can answer. There
     * MUST NOT be more than six objects in the array.
     */
    conversation_starters?: ConversationStarterElement[];
    /**
     * Optional. A list of objects that identify API plugins that provide actions accessible to
     * the declarative agent.
     */
    actions?: ActionElement[];
    [property: string]: any;
}

/**
 * Identifies an API plugin manifest for a plugin used as an action by the declarative agent.
 */
export interface ActionElement {
    /**
     * Required. Not localizable. A unique identifier for the action. It MAY be represented by a
     * GUID.
     */
    id: string;
    /**
     * Required. Not localizable. A path to the API plugin manifest for this action.
     */
    file: string;
    [property: string]: any;
}

/**
 * Represents a base capability object.
 *
 * Indicates that the declarative agent can search the web for grounding information.
 *
 * Indicates that the declarative agent can search a user's SharePoint and OneDrive for
 * grounding information.
 *
 * Indicates that the declarative agent can search selected Microsoft Graph connectors for
 * grounding information.
 *
 * Indicates that the declarative agent can generate and execute code.
 *
 * Indicates that the declarative agent can images and art based on the text input from the
 * user.
 *
 * Indicates that the declarative agent can search through Teams channels, teams, meetings,
 * 1:1 chats and group chats.
 *
 * A JSON object whose presence indicates that the DA will be able to search within Email
 * Messages in the mailboxes user has access to.
 *
 * Indicates that the DA will be able to search people data in the organization.
 */
export interface CapabilityElement {
    /**
     * Required. The name of the capability. Allowed values are WebSearch, CodeInterpreter,
     * OneDriveAndSharePoint, GraphConnectors, TeamsMessages, Dataverse, Email, People and
     * GraphicArt.
     *
     * Required. Must be set to WebSearch.
     *
     * Required. Must be set to OneDriveAndSharePoint.
     *
     * Required. Must be set to GraphConnectors.
     *
     * Required. Must be set to CodeInterpreter.
     *
     * Required. Must be set to GraphicArt.
     *
     * Required. Must be set to TeamsMessages.
     *
     * Required: Must be set to Dataverse
     *
     * Required: Must be set to Email
     *
     * Required. Must be set to People.
     */
    name: Name;
    /**
     * Optional. An array of sites used to constrain the content accessible to the DA to just
     * the content identified via the items of array.
     */
    sites?: SiteElement[];
    /**
     * Optional. An array of objects that identify SharePoint or OneDrive sources using IDs.
     */
    items_by_sharepoint_ids?: ItemsBySharepointIDElement[];
    /**
     * Optional. An array of objects that identify SharePoint or OneDrive sources by URL.
     */
    items_by_url?: ItemsByURLElement[];
    /**
     * Optional. An array of objects that identify the Microsoft Graph connectors available to
     * the declarative agent
     */
    connections?: ConnectionElement[];
    /**
     * This member can be used to constrain the content accessible to the DA to just the content
     * identified via the members of each Teams url
     */
    urls?: URLElement[];
    /**
     * An array of Objects that represent the knowledge sources for the Dataverse in the
     * Declarative Agent
     */
    knowledge_sources?: KnowledgeSourceElement[];
    /**
     * A JSON array of Folder Object. This member can be used to constrain the content
     * accessible to the DA to just the emails present in the folders identified by members of
     * each Folder Object.
     */
    folders?: FolderElement[];
    /**
     * A JSON string that contains SMTP address of the shared mailbox. The presence of this
     * field indicates that the DA constrain its search for relevant emails only to that
     * mailbox. Emails from user's primary mailbox is not searched when this field is present.
     */
    shared_mailbox?: string;
    [property: string]: any;
}

/**
 * Identifies a Microsoft Graph connector.
 */
export interface ConnectionElement {
    /**
     * Required. Not localizable The unique identifier of the Microsoft Graph connector.
     */
    connection_id: string;
    [property: string]: any;
}

export interface FolderElement {
    /**
     * A JSON string that identifies an email folder. This can either be id of the folder or one
     * of the well known names.
     */
    folder_id: string;
    [property: string]: any;
}

/**
 * Contains one or more object identifiers that identify a SharePoint or OneDrive resource.
 */
export interface ItemsBySharepointIDElement {
    /**
     * Optional. Not localizable. The GUID identifier of a SharePoint or OneDrive site.
     */
    site_id?: string;
    /**
     * Optional. Not localizable. The GUID identifier of a SharePoint or OneDrive web.
     */
    web_id?: string;
    /**
     * Optional. Not localizable. The GUID identifier of a SharePoint or OneDrive list.
     */
    list_id?: string;
    /**
     * Optional. Not localizable. The GUID identifier of a SharePoint or OneDrive item.
     */
    unique_id?: string;
    /**
     * Boolean value indicating whether to enable searching associated sites. This value is only
     * applicable when the site_id value references a SharePoint HubSite.
     */
    search_associated_sites?: boolean;
    /**
     * A JSON String that uniquely identifies a part of a SharePoint item. e.g a OneNote page.
     */
    part_id?: string;
    /**
     * A String that qualifies the kind of part that the "part_id" refers to. Currently this
     * value can only be equal to the string literal: "OneNotePart".
     */
    part_type?: "OneNotePart";
}

/**
 * A String that qualifies the kind of part that the "part_id" refers to. Currently this
 * value can only be equal to the string literal: "OneNotePart".
 */

/**
 * Represents the URL of a SharePoint or OneDrive resource.
 */
export interface ItemsByURLElement {
    /**
     * Optional. Not localizable. An absolute URL to a SharePoint or OneDrive resource.
     */
    url?: string;
}

export interface KnowledgeSourceElement {
    /**
     * A unique identifier for the host in Dataverse.
     */
    host_name?: string;
    /**
     * A unique identifier that defines the configuration for how the copilot agent interacts
     * with Dataverse knowledge.
     */
    skill?: string;
    /**
     * An array of table_name objects which contain table names in DataVerse to scope the
     * knowledge of the Declarative Agent
     */
    tables?: TableElement[];
    [property: string]: any;
}

export interface TableElement {
    /**
     * A string to represent the table name.
     */
    table_name?: string;
    [property: string]: any;
}

export type Name = "WebSearch" | "CodeInterpreter" | "OneDriveAndSharePoint" | "GraphConnectors" | "GraphicArt" | "TeamsMessages" | "Dataverse" | "Email" | "People";

/**
 * An object that identifies a site used to constrain the content accessible to the
 * declarative agent.
 */
export interface SiteElement {
    /**
     * An absolute URL to a site.
     */
    url: string;
    [property: string]: any;
}

/**
 * Identifies a Teams channel, team or meeting chat
 */
export interface URLElement {
    /**
     * A string that contains a well formed, Teams url to a Teams channel, team or meeting chat
     * (join url)
     */
    url: string;
    [property: string]: any;
}

/**
 * Contains hints that are displayed to the user to demonstrate how they can get started
 * using the declarative agent.
 */
export interface ConversationStarterElement {
    /**
     * Required. Localizable. A suggestion that the user can use to obtain the desired result
     * from the DC. It MUST contain at least one nonwhitespace character.
     */
    text: string;
    /**
     * Optional. Localizable. A unique title for the conversation starter. It MUST contain at
     * least one nonwhitespace character.
     */
    title?: string;
    [property: string]: any;
}

// Converts JSON strings to/from your types
// and asserts the results of JSON.parse at runtime
export class Convert {
    public static toDeclarativeAgentManifestV1D3(json: string): DeclarativeAgentManifestV1D3 {
        return cast(JSON.parse(json), r("DeclarativeAgentManifestV1D3"));
    }

    public static declarativeAgentManifestV1D3ToJson(value: DeclarativeAgentManifestV1D3): string {
        return JSON.stringify(uncast(value, r("DeclarativeAgentManifestV1D3")), null, 4);
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
    "DeclarativeAgentManifestV1D3": o([
        { json: "version", js: "version", typ: r("Version") },
        { json: "id", js: "id", typ: u(undefined, "") },
        { json: "name", js: "name", typ: "" },
        { json: "description", js: "description", typ: "" },
        { json: "instructions", js: "instructions", typ: u(undefined, "") },
        { json: "capabilities", js: "capabilities", typ: u(undefined, a(r("CapabilityElement"))) },
        { json: "conversation_starters", js: "conversation_starters", typ: u(undefined, a(r("ConversationStarterElement"))) },
        { json: "actions", js: "actions", typ: u(undefined, a(r("ActionElement"))) },
    ], "any"),
    "ActionElement": o([
        { json: "id", js: "id", typ: "" },
        { json: "file", js: "file", typ: "" },
    ], "any"),
    "CapabilityElement": o([
        { json: "name", js: "name", typ: r("Name") },
        { json: "sites", js: "sites", typ: u(undefined, a(r("SiteElement"))) },
        { json: "items_by_sharepoint_ids", js: "items_by_sharepoint_ids", typ: u(undefined, a(r("ItemsBySharepointIDElement"))) },
        { json: "items_by_url", js: "items_by_url", typ: u(undefined, a(r("ItemsByURLElement"))) },
        { json: "connections", js: "connections", typ: u(undefined, a(r("ConnectionElement"))) },
        { json: "urls", js: "urls", typ: u(undefined, a(r("URLElement"))) },
        { json: "knowledge_sources", js: "knowledge_sources", typ: u(undefined, a(r("KnowledgeSourceElement"))) },
        { json: "folders", js: "folders", typ: u(undefined, a(r("FolderElement"))) },
        { json: "shared_mailbox", js: "shared_mailbox", typ: u(undefined, "") },
    ], "any"),
    "ConnectionElement": o([
        { json: "connection_id", js: "connection_id", typ: "" },
    ], "any"),
    "FolderElement": o([
        { json: "folder_id", js: "folder_id", typ: "" },
    ], "any"),
    "ItemsBySharepointIDElement": o([
        { json: "site_id", js: "site_id", typ: u(undefined, "") },
        { json: "web_id", js: "web_id", typ: u(undefined, "") },
        { json: "list_id", js: "list_id", typ: u(undefined, "") },
        { json: "unique_id", js: "unique_id", typ: u(undefined, "") },
        { json: "search_associated_sites", js: "search_associated_sites", typ: u(undefined, true) },
        { json: "part_id", js: "part_id", typ: u(undefined, "") },
        { json: "part_type", js: "part_type", typ: u(undefined, r("PartType")) },
    ], false),
    "ItemsByURLElement": o([
        { json: "url", js: "url", typ: u(undefined, "") },
    ], false),
    "KnowledgeSourceElement": o([
        { json: "host_name", js: "host_name", typ: u(undefined, "") },
        { json: "skill", js: "skill", typ: u(undefined, "") },
        { json: "tables", js: "tables", typ: u(undefined, a(r("TableElement"))) },
    ], "any"),
    "TableElement": o([
        { json: "table_name", js: "table_name", typ: u(undefined, "") },
    ], "any"),
    "SiteElement": o([
        { json: "url", js: "url", typ: "" },
    ], "any"),
    "URLElement": o([
        { json: "url", js: "url", typ: "" },
    ], "any"),
    "ConversationStarterElement": o([
        { json: "text", js: "text", typ: "" },
        { json: "title", js: "title", typ: u(undefined, "") },
    ], "any"),
    "PartType": [
        "OneNotePart",
    ],
    "Name": [
        "CodeInterpreter",
        "Dataverse",
        "Email",
        "GraphConnectors",
        "GraphicArt",
        "OneDriveAndSharePoint",
        "People",
        "TeamsMessages",
        "WebSearch",
    ],
    "Version": [
        "v1.3",
    ],
};
