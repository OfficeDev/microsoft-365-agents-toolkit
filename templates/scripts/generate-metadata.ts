import fs from "fs";
import path from "path";
import { allTemplates, defaultGeneratorTemplates } from "../src/metadata/index";
import { ceaNode } from "../src/ui/cea";
import { daNode } from "../src/ui/da";
import { graphConnectorNode } from "../src/ui/graphConnector";
import { officeAddinNode } from "../src/ui/officeAddin";
import { rootNode } from "../src/ui/root";
import { teamsNode } from "../src/ui/teams";

// CLI argument parsing and execution
function main() {
  fs.mkdirSync(path.resolve(__dirname, "../build/metadata"), { recursive: true });
  fs.mkdirSync(path.resolve(__dirname, "../build/ui"), { recursive: true });

  fs.writeFileSync(
    path.resolve(__dirname, "../build/ui/rootNode.json"),
    JSON.stringify(rootNode, null, 2),
    "utf-8"
  );
  fs.writeFileSync(
    path.resolve(__dirname, "../build/ui/daNode.json"),
    JSON.stringify(daNode, null, 2),
    "utf-8"
  );
  fs.writeFileSync(
    path.resolve(__dirname, "../build/ui/ceaNode.json"),
    JSON.stringify(ceaNode, null, 2),
    "utf-8"
  );
  fs.writeFileSync(
    path.resolve(__dirname, "../build/ui/teamsNode.json"),
    JSON.stringify(teamsNode, null, 2),
    "utf-8"
  );
  fs.writeFileSync(
    path.resolve(__dirname, "../build/ui/graphConnectorNode.json"),
    JSON.stringify(graphConnectorNode, null, 2),
    "utf-8"
  );
  fs.writeFileSync(
    path.resolve(__dirname, "../build/ui/officeAddinNode.json"),
    JSON.stringify(officeAddinNode, null, 2),
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
