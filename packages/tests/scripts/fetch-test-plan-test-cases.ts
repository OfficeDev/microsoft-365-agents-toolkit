import * as azdev from "azure-devops-node-api";
import { AzureCliCredential } from "@azure/identity";
import * as fs from "fs";
import * as path from "path";

const outputDir = process.argv[2];
if (!outputDir) {
  console.error("Please provide an output directory as the first argument.");
  process.exit(1);
}
if (!fs.existsSync(outputDir)) {
  fs.mkdirSync(outputDir, { recursive: true });
}

async function run() {
  // Print environment variables for debugging
  console.log("Environment Variables:");
  console.log("AZURE_DEVOPS_ORG_URL:", process.env.AZURE_DEVOPS_ORG_URL);
  console.log("AZURE_DEVOPS_PROJECT:", process.env.AZURE_DEVOPS_PROJECT);
  console.log(
    "AZURE_DEVOPS_TEST_PLAN_ID:",
    process.env.AZURE_DEVOPS_TEST_PLAN_ID
  );
  console.log(
    "AZURE_DEVOPS_TEST_SUITE_ID:",
    process.env.AZURE_DEVOPS_TEST_SUITE_ID
  );
  const orgUrl = process.env.AZURE_DEVOPS_ORG_URL!;
  const project = process.env.AZURE_DEVOPS_PROJECT!;
  const planId = parseInt(process.env.AZURE_DEVOPS_TEST_PLAN_ID!);
  const suiteId = parseInt(process.env.AZURE_DEVOPS_TEST_SUITE_ID!);

  // Get Azure AD token using Azure CLI credential
  const credential = new AzureCliCredential();
  const token = await credential.getToken(
    "https://app.vssps.visualstudio.com/.default"
  );

  // Initialize Azure DevOps API client

  // Create auth handler using the Azure AD token
  const authHandler = azdev.getBearerHandler(token.token);
  // log the auth handler for debugging
  console.log("Auth Handler created successfully");
  const connection = new azdev.WebApi(orgUrl, authHandler);
  console.log("Connection to Azure DevOps established successfully");
  const testApi = await connection.getTestApi();
  console.log("Test API client initialized successfully");

  const testCases = await testApi.getTestCases(project, planId, suiteId);
  console.log("Test cases fetched successfully");
  console.log(
    `Found ${testCases.length} test cases in plan ${planId}, suite ${suiteId}`
  );
  for (const tc of testCases) {
    if (tc.testCase) {
      console.log(`Test case - ${tc.testCase.id}`);

      // Fetch the work item JSON from the url
      if (tc.testCase.url) {
        try {
          // Use the Azure AD token for authentication
          const response = await fetch(tc.testCase.url, {
            headers: {
              Authorization: `Bearer ${token.token}`,
              "Content-Type": "application/json",
            },
          });
          if (!response.ok) {
            console.error(
              `Failed to fetch work item: ${tc.testCase.url}, status: ${response.status}`
            );
            continue;
          }

          // Define a minimal type for workItem
          type WorkItem = {
            fields?: {
              [key: string]: any;
            };
          };
          const workItem = (await response.json()) as WorkItem;

          // The content fianlly to write into the file
          const outputLines: string[] = [];

          // Get the URL of the work item
          outputLines.push(
            `#URL: https://msazure.visualstudio.com/Microsoft%20Teams%20Extensibility/_workitems/edit/${tc.testCase.id}`
          );

          // Get the title of the work itme
          const title = workItem.fields?.["System.Title"];
          if (title) {
            outputLines.push(`#Title: ${title}`);
          }

          // Get the author of the work item
          const assignedTo = workItem.fields?.["System.AssignedTo"];
          const uniqueName =
            assignedTo && typeof assignedTo === "object"
              ? assignedTo.uniqueName
              : undefined;
          if (uniqueName) {
            outputLines.push(`#Author: ${uniqueName}`);
          }

          //if (tags && tags.includes("VSCUSE")) {
          const steps = workItem.fields?.["Microsoft.VSTS.TCM.Steps"];
          if (typeof steps === "string") {
            const stepBlocks = steps.match(/<step[\s\S]*?<\/step>/gi) || [];

            stepBlocks.forEach((stepBlock, idx) => {
              const paramMatch = stepBlock.match(
                /<parameterizedString[^>]*>([\s\S]*?)<\/parameterizedString>/i
              );
              let text = "";
              if (paramMatch && paramMatch[1]) {
                const html = paramMatch[1]
                  .replace(/&lt;/g, "<")
                  .replace(/&gt;/g, ">");
                const pMatch =
                  html.match(/<p>([\s\S]*?)<\/p>/i) ||
                  html.match(/<P>([\s\S]*?)<\/P>/i);
                if (pMatch && pMatch[1]) {
                  text = pMatch[1].replace(/<[^>]+>/g, "").trim();
                }
              }
              if (!text) {
                const match =
                  stepBlock.match(/<p>([\s\S]*?)<\/p>/i) ||
                  stepBlock.match(/<P>([\s\S]*?)<\/P>/i);
                if (match && match[1]) {
                  text = match[1].replace(/<[^>]+>/g, "").trim();
                }
              }
              if (text) {
                outputLines.push(text);
              }
            });
            const filePath = path.join(outputDir, `${tc.testCase.id}.txt`);
            fs.writeFileSync(filePath, outputLines.join("\n"), {
              encoding: "utf8",
            });
            console.log(
              `Wrote steps for test case ${tc.testCase.id} to ${filePath}`
            );
            console.log(
              `The file content is:  \n${outputLines.join(
                "\n"
              )} \n [End of file]`
            );
          } else {
            console.log(`The type is: ${typeof steps}`);
          }
          //}
        } catch (err) {
          console.error(
            `Error fetching work item for test case ${tc.testCase.id}:`,
            err
          );
        }
      }
    } else {
      console.error(`- Warning: Test case is undefined or missing details.`);
    }
  }
}

run().catch((err) => {
  console.error("Error:", err);
  process.exit(1);
});
