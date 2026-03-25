import fs from "fs";
import path from "path";
import { allTemplates, defaultGeneratorTemplates } from "../src/metadata/index";
import { tdpWizardNode, wizardNode } from "../src/ui/wizard";

function main() {
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
