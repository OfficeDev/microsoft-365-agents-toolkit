import { Inputs, IQTreeNode, OptionItem, Platform } from "@microsoft/teamsfx-api";
import * as path from "path";
import * as fs from "fs-extra";

import { TemplateNames } from "../src/component/generator/templates/templateNames";
import { questionNodes } from "../src/question";
import { QuestionNames } from "../src/question/questionNames";
import { getLanguageOptions } from "../src/question/scaffold/vsc/createRootNode";
import { getAllTemplatesOnPlatform } from "../src/component/generator/templates/metadata";

interface ITemplateContext {
  nodes: { nodeId: string; optionId: string; optionLabel: string }[];
  templateName: TemplateNames;
  templateIds: string[];
  supportedLanguages: string[];
  description: string;
  readme: string;
}

// Helper function to filter out labels containing $(xx) pattern
function filterLabel(label: string): string {
  // Remove text matching pattern $(anything)
  return label.replace(/\$\([^)]*\)\s*/g, "").trim();
}

export function traverse(node: IQTreeNode): ITemplateContext[] {
  const results: ITemplateContext[] = [];

  function traverseRecursive(
    currentNode: IQTreeNode,
    currentPath: { nodeId: string; optionId: string; optionLabel: string }[],
    pathDetails: string[] = []
  ): void {
    // Skip group nodes, just traverse their children
    if (currentNode.data.type === "group") {
      if (currentNode.children) {
        for (const child of currentNode.children) {
          if (child) {
            traverseRecursive(child, currentPath, pathDetails);
          }
        }
      }
      return;
    }

    // Handle question nodes
    const question = currentNode.data as {
      name?: string;
      staticOptions?: OptionItem[];
      type: string;
      skipSingleOption?: boolean;
    };

    // Only process selection nodes (singleSelect)
    if (question.type !== "singleSelect") {
      return;
    }

    // Process selection nodes
    if (question.staticOptions && Array.isArray(question.staticOptions)) {
      // Check if we should skip recording this selection
      const shouldSkipRecording = question.staticOptions.length === 1 && question.skipSingleOption;

      for (const option of question.staticOptions) {
        if (option && typeof option === "object") {
          // Check if this option has a data field that is a TemplateNames
          if (
            "data" in option &&
            option.data &&
            Object.values(TemplateNames).includes(option.data as TemplateNames)
          ) {
            // Create path with this option selection (skip if single option and skipSingleOption is true)
            const pathWithCurrentSelection: {
              nodeId: string;
              optionId: string;
              optionLabel: string;
            }[] =
              shouldSkipRecording || !question.name
                ? currentPath
                : [
                    ...currentPath,
                    {
                      nodeId: question.name,
                      optionId: option.id,
                      optionLabel: filterLabel(option.label),
                    },
                  ];

            // Create path details with this option's detail (skip if single option and skipSingleOption is true)
            const pathDetailsWithCurrent =
              shouldSkipRecording || !option.detail ? pathDetails : [...pathDetails, option.detail];

            // Create template context with joined descriptions
            const templateContext: ITemplateContext = {
              nodes: pathWithCurrentSelection,
              templateName: option.data as TemplateNames,
              description:
                pathDetailsWithCurrent.length > 0
                  ? pathDetailsWithCurrent.join(" - ")
                  : option.detail || option.label,
              readme: "",
              supportedLanguages: [], // Initialize with an empty array
              templateIds: [], // Initialize with an empty array
            };

            results.push(templateContext);
            continue;
          }

          // Traverse children that match this option selection
          if (currentNode.children) {
            const pathWithCurrentSelection: {
              nodeId: string;
              optionId: string;
              optionLabel: string;
            }[] =
              shouldSkipRecording || !question.name
                ? currentPath
                : [
                    ...currentPath,
                    {
                      nodeId: question.name,
                      optionId: option.id,
                      optionLabel: filterLabel(option.label),
                    },
                  ];

            // Create path details with this option's detail (skip if single option and skipSingleOption is true)
            const pathDetailsWithCurrent =
              shouldSkipRecording || !option.detail ? pathDetails : [...pathDetails, option.detail];

            for (const child of currentNode.children) {
              if (child && shouldChildBeActive(child, option.id)) {
                traverseRecursive(child, pathWithCurrentSelection, pathDetailsWithCurrent);
              }
            }
          }
        }
      }
    }
  }

  // Helper function to check if a child node should be active for a given parent value
  function shouldChildBeActive(child: IQTreeNode, parentValue: string): boolean {
    if (!child.condition) {
      return true; // No condition means always active
    }

    // Handle simple string equality conditions like { equals: "option-id" }
    const condition = child.condition as { equals?: string };
    if (condition.equals) {
      return condition.equals === parentValue;
    }

    // Handle other condition types if needed
    // For now, assume true for other condition types
    return true;
  }

  traverseRecursive(node, [], []);
  return results;
}

function setSupportedLanguagesAndTemplateIds(
  results: ITemplateContext[],
  platform: Platform
): ITemplateContext[] {
  for (const context of results) {
    const inputs = {
      [QuestionNames.TemplateName]: context.templateName,
      platform: platform,
    } as Inputs;
    const options = getLanguageOptions(inputs);
    context.supportedLanguages = options.map((opt) => opt.id);
    context.templateIds = getAllTemplatesOnPlatform(inputs.platform)
      .filter((t) => t.name === context.templateName)
      .filter((t) => t.language && context.supportedLanguages.includes(t.language))
      .map((t) => t.id);
  }
  return results;
}

function setReadme(results: ITemplateContext[], platform: Platform): ITemplateContext[] {
  for (const context of results) {
    // Map platform to directory name
    const platformDir = platform.toLocaleLowerCase();

    for (const templateId of context.templateIds) {
      let languageDir = templateId.substring(templateId.lastIndexOf("-") + 1);
      let templateFolderName = templateId.substring(0, templateId.lastIndexOf("-"));

      if (
        !templateId.endsWith("-ts") &&
        !templateId.endsWith("-js") &&
        !templateId.endsWith("-python") &&
        !templateId.endsWith("-csharp")
      ) {
        languageDir = "common";
        templateFolderName = templateId;
      }

      // Construct the template path
      const templatePath = path.join(
        __dirname,
        "..",
        "..",
        "..",
        "templates",
        platformDir,
        languageDir,
        templateFolderName
      );

      // Try to read README.md.tpl first, then README.md
      const readmeFiles = ["README.md.tpl", "README.md"];

      for (const readmeFile of readmeFiles) {
        const readmePath = path.join(templatePath, readmeFile);
        try {
          if (fs.existsSync(readmePath)) {
            context.readme = fs.readFileSync(readmePath, "utf-8");
            break;
          }
        } catch (error) {
          // Continue to next file if this one fails to read
          continue;
        }
      }
      if (context.readme) {
        break; // Only need to find readme for the first valid templateId
      }
    }
  }

  return results;
}
export function generateTemplateContexts(platform: Platform): ITemplateContext[] {
  // Skip CLI platform
  if (platform === Platform.CLI) {
    console.log("CLI platform is ignored, returning empty results.");
    return [];
  }

  // Import question nodes here to avoid circular dependency
  const node = questionNodes["createProject"](platform);
  const results = traverse(node);
  setSupportedLanguagesAndTemplateIds(results, platform);
  setReadme(results, platform);
  return results;
}

// CLI argument parsing and execution
function main() {
  const args = process.argv.slice(2);

  if (args.length < 2) {
    console.error("Usage: npm run generate-template <platform> <output-file-path>");
    console.error("Platforms: vscode, vs");
    console.error("Example: npm run generate-template vscode ./output/templates.json");
    process.exit(1);
  }

  const platformArg = args[0].toLowerCase();
  const outputPath = args[1];

  // Map platform string to Platform enum
  let platform: Platform;
  switch (platformArg) {
    case "vscode":
      platform = Platform.VSCode;
      break;
    case "vs":
      platform = Platform.VS;
      break;
    default:
      console.error(`Invalid platform: ${platformArg}`);
      console.error("Valid platforms: vscode, vs");
      process.exit(1);
  }

  try {
    console.log(`Generating template contexts for platform: ${platformArg}`);
    const results = generateTemplateContexts(platform);

    // Ensure output directory exists
    const outputDir = path.dirname(outputPath);
    fs.ensureDirSync(outputDir);

    // Write results to file
    fs.writeFileSync(outputPath, JSON.stringify(results, null, 2), "utf-8");

    console.log(`Template contexts generated successfully!`);
    console.log(`Output written to: ${outputPath}`);
    console.log(`Total templates found: ${results.length}`);
  } catch (error) {
    console.error("Error generating template contexts:", error);
    process.exit(1);
  }
}

// Run the script if called directly
if (require.main === module) {
  main();
}
