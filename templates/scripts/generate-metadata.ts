import fs from "fs";
import path from "path";
import { allTemplates, defaultGeneratorTemplates } from "../src/metadata/index";
import { tdpWizardNode, wizardNode } from "../src/ui/wizard";

/** Fields whose string values must be NLS keys (template.* or core.*). */
const NLS_FIELDS = new Set(["label", "detail", "title", "placeholder", "groupName"]);

/**
 * Recursively walk a wizard node tree and assert that all user-visible string
 * fields are proper NLS keys.  Raw English strings (e.g. "None", "API Key")
 * would be passed verbatim to getLocalizedString() at runtime and return "",
 * causing blank display text.
 */
function assertNoHardcodedStrings(node: unknown, path: string): void {
  if (node === null || typeof node !== "object") return;
  if (Array.isArray(node)) {
    node.forEach((item, i) => assertNoHardcodedStrings(item, `${path}[${i}]`));
    return;
  }
  for (const [key, value] of Object.entries(node as Record<string, unknown>)) {
    if (NLS_FIELDS.has(key) && typeof value === "string") {
      if (!value.startsWith("template.") && !value.startsWith("core.")) {
        throw new Error(
          `Hardcoded string found at ${path}.${key}: "${value}"\n` +
            `  All label/detail/title/placeholder/groupName values must be NLS keys ` +
            `(starting with "template." or "core.").`
        );
      }
    } else {
      assertNoHardcodedStrings(value, `${path}.${key}`);
    }
  }
}

function main() {
  // Validate that no hardcoded strings slipped through before writing JSON.
  assertNoHardcodedStrings(wizardNode, "wizardNode");
  assertNoHardcodedStrings(tdpWizardNode, "tdpWizardNode");

  fs.mkdirSync(path.resolve(__dirname, "../build/metadata"), { recursive: true });
  fs.mkdirSync(path.resolve(__dirname, "../build/ui"), { recursive: true });

  fs.writeFileSync(
    path.resolve(__dirname, "../build/ui/wizardNode.json"),
    JSON.stringify(wizardNode, null, 2),
    "utf-8"
  );
  fs.writeFileSync(
    path.resolve(__dirname, "../build/ui/tdpNode.json"),
    JSON.stringify(tdpWizardNode, null, 2),
    "utf-8"
  );
  fs.writeFileSync(
    path.resolve(__dirname, "../build/metadata/allTemplates.json"),
    JSON.stringify(allTemplates, null, 2),
    "utf-8"
  );
  fs.writeFileSync(
    path.resolve(__dirname, "../build/metadata/defaultGeneratorTemplates.json"),
    JSON.stringify(defaultGeneratorTemplates, null, 2),
    "utf-8"
  );
}

// Run the script if called directly
if (require.main === module) {
  main();
}
