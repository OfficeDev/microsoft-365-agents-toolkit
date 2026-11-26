// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

// The SPFx build leaves `{{TENANT_ORIGIN}}` in the content security policy that the generated MCP
// tools declare for their UI. Only those two lists are qualified for resolution. Any other
// unresolved whole-value placeholder is refused instead of being published literally.

const tenantOriginToken = "{{TENANT_ORIGIN}}";
const tenantOriginPlaceholder = /^\{\{\s*TENANT_ORIGIN\s*\}\}$/;
const placeholderPattern = /\{\{\s*([A-Z][A-Z0-9_]*)\s*\}\}/;
const wholeValuePlaceholder = /^\{\{\s*([A-Z][A-Z0-9_]*)\s*\}\}$/;
const qualifiedContentSecurityLists = ["frameDomains", "resourceDomains"];
// Generated plugins nest about a dozen levels. The bound keeps the indented serialization of the
// resolved plugin far below the call stack limit, which JSON.stringify reaches near 5000 levels.
const maxPluginDepth = 128;

export type UnresolvedTenantValue =
  { readonly kind: "tooDeep" } | { readonly kind: "placeholder"; readonly token: string };

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function qualifiedTenantOriginLists(runtime: Record<string, unknown>): unknown[][] {
  const spec = runtime["spec"];
  const description = isRecord(spec) ? spec["mcp_tool_description"] : undefined;
  const tools = isRecord(description) ? description["tools"] : undefined;
  if (!Array.isArray(tools)) {
    return [];
  }
  const lists: unknown[][] = [];
  for (const tool of tools) {
    const metadata = isRecord(tool) ? tool["_meta"] : undefined;
    const policy = isRecord(metadata) ? metadata["ui/csp"] : undefined;
    if (!isRecord(policy)) {
      continue;
    }
    for (const field of qualifiedContentSecurityLists) {
      const list = policy[field];
      if (Array.isArray(list)) {
        lists.push(list);
      }
    }
  }
  return lists;
}

/**
 * Resolves the tenant origin placeholder in the qualified content security policy lists of a
 * tenant-bound runtime. Returns false when a qualified entry carries any other placeholder.
 */
export function resolveTenantOrigin(runtime: Record<string, unknown>, origin: string): boolean {
  for (const list of qualifiedTenantOriginLists(runtime)) {
    for (let index = 0; index < list.length; index += 1) {
      const entry = list[index];
      if (typeof entry !== "string") {
        continue;
      }
      if (tenantOriginPlaceholder.test(entry.trim())) {
        list[index] = origin;
      } else if (placeholderPattern.test(entry)) {
        return false;
      }
    }
  }
  return true;
}

/** Undoes resolveTenantOrigin so a resolved plugin compares equal to its tenant-neutral source. */
export function restoreTenantOrigin(runtime: Record<string, unknown>, origin: string): void {
  for (const list of qualifiedTenantOriginLists(runtime)) {
    for (let index = 0; index < list.length; index += 1) {
      if (list[index] === origin) {
        list[index] = tenantOriginToken;
      }
    }
  }
}

/**
 * Finds the first string that is entirely a placeholder, or reports a plugin nested too deeply to
 * serialize safely. Partial placeholders inside free text are preserved and never reported. The
 * traversal is iterative so that hostile nesting cannot overflow the call stack.
 */
export function findUnresolvedTenantValue(plugin: unknown): UnresolvedTenantValue | undefined {
  const pending: { value: unknown; depth: number }[] = [{ value: plugin, depth: 0 }];
  for (let next = pending.pop(); next !== undefined; next = pending.pop()) {
    const { value, depth } = next;
    if (typeof value === "string") {
      const token = wholeValuePlaceholder.exec(value.trim())?.[1];
      if (token !== undefined) {
        return { kind: "placeholder", token };
      }
      continue;
    }
    let children: unknown[];
    if (Array.isArray(value)) {
      children = value;
    } else if (isRecord(value)) {
      children = Object.values(value);
    } else {
      continue;
    }
    if (depth >= maxPluginDepth) {
      return { kind: "tooDeep" };
    }
    for (let index = children.length - 1; index >= 0; index -= 1) {
      pending.push({ value: children[index], depth: depth + 1 });
    }
  }
  return undefined;
}
