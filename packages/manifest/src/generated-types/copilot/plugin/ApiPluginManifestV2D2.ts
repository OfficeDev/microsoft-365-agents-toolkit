// To parse this data:
//
//   import { Convert, APIPluginManifestV2D2 } from "./file";
//
//   const aPIPluginManifestV2D2 = Convert.toAPIPluginManifestV2D2(json);
//
// These functions will throw an error if the JSON doesn't
// match the expected interface, even if the JSON is valid.

/**
 * The root of the plugin manifest document is a JSON object that contains members that
 * describe the plugin.
 */
export interface APIPluginManifestV2D2 {
    /**
     * The schema version. Previous versions are `v1` and `v2`, `v2.1`.
     */
    schema_version: "v2.2";
    /**
     * A short, human-readable name for the plugin. It MUST contain at least one nonwhitespace
     * character. Characters beyond 20 MAY be ignored. This property is localizable.
     */
    name_for_human: string;
    /**
     * An identifier used to prevent name conflicts between function names from different
     * plugins that are used within the same execution context. The value MUST match the regex
     * ^[A-Za-z0-9_]+ as defined by [RFC9485]. This is a required member.
     */
    namespace: string;
    /**
     * The description for the plugin that is provided to the model. This description should
     * describe what the plugin is for, and in what circumstances its functions are relevant.
     * Characters beyond 2048 MAY be ignored. This property is localizable.
     */
    description_for_model?: string;
    /**
     * A human-readable description of the plugin. Characters beyond 100 MAY be ignored. This
     * property is localizable.
     */
    description_for_human: string;
    /**
     * A URL used to fetch a logo that MAY be used by the orchestrator. Implementations MAY
     * provide alternative methods to provide logos that meet their visual requirements. This
     * property is localizable.
     */
    logo_url?: any;
    /**
     * An email address of a contact for safety/moderation, support, and deactivation.
     */
    contact_email?: string;
    /**
     * An absolute URL that locates a document containing the terms of service for the plugin.
     * This property is localizable.
     */
    legal_info_url?: any;
    /**
     * An absolute URL that locates a document containing the privacy policy for the plugin.
     * This property is localizable.
     */
    privacy_policy_url?: any;
    /**
     * A set of function objects describing the functions available to the plugin. Each function
     * object name MUST be unique within the array. The order of the array isn't significant. If
     * the `functions` property isn't present and there's an OpenAPI runtime, the functions are
     * inferred from the OpenAPI operations.
     */
    functions?: FunctionObject[];
    /**
     * A set of runtime objects describing the runtimes used by the plugin.
     */
    runtimes?: APIPluginManifestV2D[];
    /**
     * Describes capabilities of the plugin.
     */
    capabilities?: PluginCapabilitiesObject;
    [property: string]: any;
}

/**
 * Describes capabilities of the plugin.
 */
export interface PluginCapabilitiesObject {
    /**
     * Conversation starters that can be displayed to the user for suggestions on how to invoke
     * the plugin.
     */
    conversation_starters?: ConversationStarterObject[];
    [property: string]: any;
}

/**
 * An example of a question that the plugin can answer.
 */
export interface ConversationStarterObject {
    /**
     * The text of the conversation starter. This property is localizable.
     */
    text: string;
    /**
     * The title of the conversation starter. This property is localizable.
     */
    title?: string;
    [property: string]: any;
}

/**
 * Information related to how the model should interact with a function.
 */
export interface FunctionObject {
    id?: string;
    /**
     * A string that uniquely identifies this function. Runtime objects MAY reference this
     * identifier to bind the runtime to the function. When the function is bound to an OpenAPI
     * runtime, the value must match an `operationId` value in the OpenAPI description.
     */
    name: string;
    /**
     * A description better tailored to the model, such as token context length considerations
     * or keyword usage for improved plugin prompting.
     */
    description?: string;
    /**
     * An object that contains members that describe the parameters of a function in a runtime
     * agnostic way. It mirrors the shape of [json-schema][] but only supports a small subset of
     * the JSON schema capabilities. If the `parameters` property isn't present, functions
     * described by a runtime object of type `OpenApi` use the OpenAPI description to determine
     * the parameters. Each member in the JSON object is a function parameter object that
     * describes the semantics of the parameter.
     */
    parameters?: FunctionParametersObject;
    /**
     * Describes the semantics of the value returned from the function.
     */
    returns?: ReturnObject;
    /**
     * Defines state objects for orchestrator states.
     */
    states?: FunctionStatesObject;
    /**
     * Contains a collection of data used to configure optional capabilities of the orchestrator
     * while invoking the function.
     */
    capabilities?: FunctionCapabilitiesObject;
    [property: string]: any;
}

/**
 * Contains a collection of data used to configure optional capabilities of the orchestrator
 * while invoking the function.
 */
export interface FunctionCapabilitiesObject {
    /**
     * Describes a confirmation dialog that SHOULD be presented to the user before invoking the
     * function.
     */
    confirmation?: ConfirmationObject;
    /**
     * Describes how the orchestrator can interpret the response payload and provide a visual
     * rendering.
     */
    response_semantics?: ResponseSemanticsObject;
    /**
     * Describes the security information to be used to aid in determining the relative risk of
     * invoking the function.
     */
    security_info?: SecurityInfoObject;
    [property: string]: any;
}

/**
 * Describes a confirmation dialog that SHOULD be presented to the user before invoking the
 * function.
 *
 * Describes how the orchestrator asks the user to confirm before calling a function.
 */
export interface ConfirmationObject {
    /**
     * Specifies the type of confirmation.
     */
    type?: ConfirmationType;
    /**
     * The title of the confirmation dialog. This property is localizable.
     */
    title?: string;
    /**
     * The text of the confirmation dialog. This property is localizable.
     */
    body?: string;
    [property: string]: any;
}

/**
 * Specifies the type of confirmation.
 */
export type ConfirmationType = "None" | "AdaptiveCard";

/**
 * Describes how the orchestrator can interpret the response payload and provide a visual
 * rendering.
 *
 * Contains information to identify semantics of response payload and enable rendering that
 * information in a rich visual experience using [adaptive cards](https://adaptivecards.io/).
 */
export interface ResponseSemanticsObject {
    /**
     * A JSONPath [RFC9535][] query that identifies a set of elements from the function response
     * to be rendered using the template specified in each item.
     */
    data_path: string;
    /**
     * Allows mapping of JSONPath queries to well-known data elements. Each JSONPath query is
     * relative to a result value.
     */
    properties?: ResponseSemanticsPropertiesObject;
    /**
     * A JSON object that conforms with the [Adaptive Card
     * Schema](https://adaptivecards.io/schemas/adaptive-card.json) and templating language.
     * This Adaptive Card instance is used to render a result from the plugin response. This
     * value is used if the `template_selector` isn't present or fails to resolve to an adaptive
     * card.
     */
    static_template?: { [key: string]: any };
    /**
     * A JSON string containing a JSONPath query that when applied to the response payload will
     * return an [Adaptive Card
     * Template](https://learn.microsoft.com/adaptive-cards/templating/language) that will be
     * used to authenticate the user.
     */
    oauth_card_path?: string;
    [property: string]: any;
}

/**
 * Allows mapping of JSONPath queries to well-known data elements. Each JSONPath query is
 * relative to a result value.
 */
export interface ResponseSemanticsPropertiesObject {
    /**
     * Title of a citation for the result.
     */
    title?: string;
    /**
     * Subtitle of a citation for the result.
     */
    subtitle?: string;
    /**
     * URL of a citation for the result.
     */
    url?: string;
    /**
     * URL of a thumbnail image for the result.
     */
    thumbnail_url?: string;
    /**
     * Data sensitivity indicator of the result contents.
     */
    information_protection_label?: string;
    /**
     * A JSONPath query that returns an [Adaptive Card
     * Template](https://learn.microsoft.com/adaptive-cards/templating/language) from the API
     * response to be used for rendering the result.
     */
    template_selector?: string;
    [property: string]: any;
}

/**
 * Describes the security information to be used to aid in determining the relative risk of
 * invoking the function.
 */
export interface SecurityInfoObject {
    /**
     * An array of strings that describe the data handling behavior of the plugin.
     */
    data_handling?: DataHandling[];
    [property: string]: any;
}

export type DataHandling = "GetPublicData" | "GetPrivateData" | "DataTransform" | "ResourceStateUpdate";

/**
 * An object that contains members that describe the parameters of a function in a runtime
 * agnostic way. It mirrors the shape of [json-schema][] but only supports a small subset of
 * the JSON schema capabilities. If the `parameters` property isn't present, functions
 * described by a runtime object of type `OpenApi` use the OpenAPI description to determine
 * the parameters. Each member in the JSON object is a function parameter object that
 * describes the semantics of the parameter.
 *
 * An object that is used to identify the set of parameters that can be passed to the
 * function. This object is structured to mirror the shape of a JSON Schema object but it
 * only supports a subset of JSON Schema keywords.
 */
export interface FunctionParametersObject {
    /**
     * The JSON Schema type.
     */
    type?: "object";
    /**
     * An object that maps parameter names to their definitions.
     */
    properties: { [key: string]: any };
    /**
     * The names of properties that are required parameters. Unlike in JSON Schema, the values
     * in this array MUST match the names listed in the `properties` property.
     */
    required?: string[];
    [property: string]: any;
}

/**
 * Describes the semantics of the value returned from the function.
 *
 * Contains the semantics of the value returned from the function.
 *
 * Indicates that the function returns a response that is compatible with the Rich Responses
 * protocol.
 */
export interface ReturnObject {
    /**
     * Specifies the type of the value returned by the API.
     */
    type?: "string";
    /**
     * A description of the value returned by the API.
     */
    description?: string;
    $ref?:        "https://copilot.microsoft.com/schemas/rich-response-v1.0.json";
    [property: string]: any;
}

/**
 * Specifies the type of the value returned by the API.
 */

/**
 * Defines state objects for orchestrator states.
 */
export interface FunctionStatesObject {
    /**
     * The state in which the model can call functions and do computations.
     */
    reasoning?: StateObject;
    /**
     * The state in which the model can generate text that is shown to the user. The model can't
     * invoke functions in the responding state.
     */
    responding?: StateObject;
    [property: string]: any;
}

/**
 * The state in which the model can call functions and do computations.
 *
 * Contains specific instructions for when a function is invoked in a specific orchestrator
 * state.
 *
 * The state in which the model can generate text that is shown to the user. The model can't
 * invoke functions in the responding state.
 */
export interface StateObject {
    /**
     * Describes the purpose of a function when used in a specific orchestrator state.
     */
    description?: string;
    /**
     * A string or an array of strings that are used to provide instructions to the orchestrator
     * on how to use this function while in a specific orchestrator state. Providing a single
     * string indicates the intent to provide a complete set of instructions that would override
     * any built-in function prompts. Providing an array of strings indicates the intent to
     * augment the built-in function prompting mechanism.
     */
    instructions?: string[] | string;
    /**
     * A string or an array of strings that are used to provide examples to the orchestrator on
     * how this function can be invoked.
     */
    examples?: string[] | string;
    [property: string]: any;
}

/**
 * A JSON object that describes the mechanics of how a function will be invoked. More than
 * one runtime MUST NOT declare support for the same function either implicitly or
 * explicitly using `run_for_functions`.
 */
export interface APIPluginManifestV2D {
    /**
     * The type of runtime. Must be 'OpenApi' or 'LocalPlugin'.
     */
    type: RuntimeType;
    auth: Auth;
    /**
     * The names of the functions that are available in this runtime. If this property is
     * omitted, all functions described by the runtime are available. If a wildcard ("*") is
     * specified as the only string, all functions are considered. More than one runtime MUST
     * NOT declare support for the same function either implicitly or explicitly.
     */
    run_for_functions?: string[];
    /**
     * Runtime-specific configuration object.
     */
    spec: Spec;
    /**
     * A Liquid template used to transform the plugin response payload.
     */
    output_template?: string;
}

/**
 * Authentication configuration required to invoke the runtime.
 */
export interface Auth {
    /**
     * Specifies the type of authentication required to invoke a function.
     */
    type: TypeEnum;
    /**
     * Specifies the type of authentication required to invoke a function.
     */
    Type?: TypeEnum;
    /**
     * A value used when `type` is `OAuthPluginVault` or `ApiKeyPluginVault`. The `reference_id`
     * value is acquired independently when providing the necessary authentication configuration
     * values. This mechanism exists to prevent the need for storing secret values in the plugin
     * manifest.
     */
    reference_id?: string;
}

/**
 * Specifies the type of authentication required to invoke a function.
 */
export type TypeEnum = "None" | "OAuthPluginVault" | "ApiKeyPluginVault";

/**
 * Runtime-specific configuration object.
 *
 * Configuration for invoking an OpenAPI-based runtime.
 *
 * Configuration for invoking a local plugin runtime.
 */
export interface Spec {
    /**
     * The URL to fetch the OpenAPI specification, called with a GET request. This member is
     * required unless `api_description` is present
     */
    url?: string;
    /**
     * A string that contains an OpenAPI description. If this member is present, `url` isn't
     * required and is ignored if present.
     */
    api_description?: string;
    /**
     * A JSON string that contains the progress style that will be used to display the progress
     * of the function. The value MUST be one of the following values: None, ShowUsage,
     * ShowUsageWithInput, ShowUsageWithInputAndOutput.
     */
    progress_style?: ProgressStyle;
    /**
     * A JSON string that represents a local runtime identifier that links to a specific
     * function to invoke locally (e.g. in the case of Windows it will link to a particular
     * app). In the case of an Office Addin that is implementing the function, the value MUST be
     * the string Microsoft.Office.Addin.
     */
    local_endpoint?: "Microsoft.Office.Addin";
}

/**
 * A JSON string that represents a local runtime identifier that links to a specific
 * function to invoke locally (e.g. in the case of Windows it will link to a particular
 * app). In the case of an Office Addin that is implementing the function, the value MUST be
 * the string Microsoft.Office.Addin.
 */

/**
 * A JSON string that contains the progress style that will be used to display the progress
 * of the function. The value MUST be one of the following values: None, ShowUsage,
 * ShowUsageWithInput, ShowUsageWithInputAndOutput.
 */
export type ProgressStyle = "None" | "ShowUsage" | "ShowUsageWithInput" | "ShowUsageWithInputAndOutput";

/**
 * The type of runtime. Must be 'OpenApi' or 'LocalPlugin'.
 */
export type RuntimeType = "OpenApi" | "LocalPlugin";

// Converts JSON strings to/from your types
// and asserts the results of JSON.parse at runtime
export class Convert {
    public static toAPIPluginManifestV2D2(json: string): APIPluginManifestV2D2 {
        return cast(JSON.parse(json), r("APIPluginManifestV2D2"));
    }

    public static aPIPluginManifestV2D2ToJson(value: APIPluginManifestV2D2): string {
        return JSON.stringify(uncast(value, r("APIPluginManifestV2D2")), null, 4);
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
    "APIPluginManifestV2D2": o([
        { json: "schema_version", js: "schema_version", typ: r("SchemaVersion") },
        { json: "name_for_human", js: "name_for_human", typ: "" },
        { json: "namespace", js: "namespace", typ: "" },
        { json: "description_for_model", js: "description_for_model", typ: u(undefined, "") },
        { json: "description_for_human", js: "description_for_human", typ: "" },
        { json: "logo_url", js: "logo_url", typ: u(undefined, "any") },
        { json: "contact_email", js: "contact_email", typ: u(undefined, "") },
        { json: "legal_info_url", js: "legal_info_url", typ: u(undefined, "any") },
        { json: "privacy_policy_url", js: "privacy_policy_url", typ: u(undefined, "any") },
        { json: "functions", js: "functions", typ: u(undefined, a(r("FunctionObject"))) },
        { json: "runtimes", js: "runtimes", typ: u(undefined, a(r("APIPluginManifestV2D"))) },
        { json: "capabilities", js: "capabilities", typ: u(undefined, r("PluginCapabilitiesObject")) },
    ], "any"),
    "PluginCapabilitiesObject": o([
        { json: "conversation_starters", js: "conversation_starters", typ: u(undefined, a(r("ConversationStarterObject"))) },
    ], "any"),
    "ConversationStarterObject": o([
        { json: "text", js: "text", typ: "" },
        { json: "title", js: "title", typ: u(undefined, "") },
    ], "any"),
    "FunctionObject": o([
        { json: "id", js: "id", typ: u(undefined, "") },
        { json: "name", js: "name", typ: "" },
        { json: "description", js: "description", typ: u(undefined, "") },
        { json: "parameters", js: "parameters", typ: u(undefined, r("FunctionParametersObject")) },
        { json: "returns", js: "returns", typ: u(undefined, r("ReturnObject")) },
        { json: "states", js: "states", typ: u(undefined, r("FunctionStatesObject")) },
        { json: "capabilities", js: "capabilities", typ: u(undefined, r("FunctionCapabilitiesObject")) },
    ], "any"),
    "FunctionCapabilitiesObject": o([
        { json: "confirmation", js: "confirmation", typ: u(undefined, r("ConfirmationObject")) },
        { json: "response_semantics", js: "response_semantics", typ: u(undefined, r("ResponseSemanticsObject")) },
        { json: "security_info", js: "security_info", typ: u(undefined, r("SecurityInfoObject")) },
    ], "any"),
    "ConfirmationObject": o([
        { json: "type", js: "type", typ: u(undefined, r("ConfirmationType")) },
        { json: "title", js: "title", typ: u(undefined, "") },
        { json: "body", js: "body", typ: u(undefined, "") },
    ], "any"),
    "ResponseSemanticsObject": o([
        { json: "data_path", js: "data_path", typ: "" },
        { json: "properties", js: "properties", typ: u(undefined, r("ResponseSemanticsPropertiesObject")) },
        { json: "static_template", js: "static_template", typ: u(undefined, m("any")) },
        { json: "oauth_card_path", js: "oauth_card_path", typ: u(undefined, "") },
    ], "any"),
    "ResponseSemanticsPropertiesObject": o([
        { json: "title", js: "title", typ: u(undefined, "") },
        { json: "subtitle", js: "subtitle", typ: u(undefined, "") },
        { json: "url", js: "url", typ: u(undefined, "") },
        { json: "thumbnail_url", js: "thumbnail_url", typ: u(undefined, "") },
        { json: "information_protection_label", js: "information_protection_label", typ: u(undefined, "") },
        { json: "template_selector", js: "template_selector", typ: u(undefined, "") },
    ], "any"),
    "SecurityInfoObject": o([
        { json: "data_handling", js: "data_handling", typ: u(undefined, a(r("DataHandling"))) },
    ], "any"),
    "FunctionParametersObject": o([
        { json: "type", js: "type", typ: u(undefined, r("ParametersType")) },
        { json: "properties", js: "properties", typ: m("any") },
        { json: "required", js: "required", typ: u(undefined, a("")) },
    ], "any"),
    "ReturnObject": o([
        { json: "type", js: "type", typ: u(undefined, r("ReturnsType")) },
        { json: "description", js: "description", typ: u(undefined, "") },
        { json: "$ref", js: "$ref", typ: u(undefined, r("Ref")) },
    ], "any"),
    "FunctionStatesObject": o([
        { json: "reasoning", js: "reasoning", typ: u(undefined, r("StateObject")) },
        { json: "responding", js: "responding", typ: u(undefined, r("StateObject")) },
    ], "any"),
    "StateObject": o([
        { json: "description", js: "description", typ: u(undefined, "") },
        { json: "instructions", js: "instructions", typ: u(undefined, u(a(""), "")) },
        { json: "examples", js: "examples", typ: u(undefined, u(a(""), "")) },
    ], "any"),
    "APIPluginManifestV2D": o([
        { json: "type", js: "type", typ: r("RuntimeType") },
        { json: "auth", js: "auth", typ: r("Auth") },
        { json: "run_for_functions", js: "run_for_functions", typ: u(undefined, a("")) },
        { json: "spec", js: "spec", typ: r("Spec") },
        { json: "output_template", js: "output_template", typ: u(undefined, "") },
    ], false),
    "Auth": o([
        { json: "type", js: "type", typ: r("TypeEnum") },
        { json: "Type", js: "Type", typ: u(undefined, r("TypeEnum")) },
        { json: "reference_id", js: "reference_id", typ: u(undefined, "") },
    ], false),
    "Spec": o([
        { json: "url", js: "url", typ: u(undefined, "") },
        { json: "api_description", js: "api_description", typ: u(undefined, "") },
        { json: "progress_style", js: "progress_style", typ: u(undefined, r("ProgressStyle")) },
        { json: "local_endpoint", js: "local_endpoint", typ: u(undefined, r("LocalEndpoint")) },
    ], false),
    "ConfirmationType": [
        "AdaptiveCard",
        "None",
    ],
    "DataHandling": [
        "DataTransform",
        "GetPrivateData",
        "GetPublicData",
        "ResourceStateUpdate",
    ],
    "ParametersType": [
        "object",
    ],
    "Ref": [
        "https://copilot.microsoft.com/schemas/rich-response-v1.0.json",
    ],
    "ReturnsType": [
        "string",
    ],
    "TypeEnum": [
        "ApiKeyPluginVault",
        "None",
        "OAuthPluginVault",
    ],
    "LocalEndpoint": [
        "Microsoft.Office.Addin",
    ],
    "ProgressStyle": [
        "None",
        "ShowUsage",
        "ShowUsageWithInput",
        "ShowUsageWithInputAndOutput",
    ],
    "RuntimeType": [
        "LocalPlugin",
        "OpenApi",
    ],
    "SchemaVersion": [
        "v2.2",
    ],
};
