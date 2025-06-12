import * as azdev from "azure-devops-node-api";
import { AzureCliCredential } from "@azure/identity";
import fetch from "node-fetch"; // Add this import at the top

async function run() {
  // Print environment variables for debugging
  console.log("Environment Variables:");
  console.log("AZURE_DEVOPS_ORG_URL:", process.env.AZURE_DEVOPS_ORG_URL);
  console.log("AZURE_DEVOPS_PROJECT:", process.env.AZURE_DEVOPS_PROJECT);
  console.log("AZURE_DEVOPS_TEST_PLAN_ID:", process.env.AZURE_DEVOPS_TEST_PLAN_ID);
  console.log("AZURE_DEVOPS_TEST_SUITE_ID:", process.env.AZURE_DEVOPS_TEST_SUITE_ID);
  const orgUrl = process.env.AZURE_DEVOPS_ORG_URL!;
  const project = process.env.AZURE_DEVOPS_PROJECT!;
  const planId = parseInt(process.env.AZURE_DEVOPS_TEST_PLAN_ID!);
  const suiteId = parseInt(process.env.AZURE_DEVOPS_TEST_SUITE_ID!);


  // Get Azure AD token using Azure CLI credential
  const credential = new AzureCliCredential();
  const token = await credential.getToken("https://app.vssps.visualstudio.com/.default");
  // Log the token for debugging (optional, be cautious with sensitive data)
  console.log("Azure AD Token GOT success, its length is:", token.token.length);
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
  console.log(`Found ${testCases.length} test cases in plan ${planId}, suite ${suiteId}`);
  for (const tc of testCases) {
    if (tc.testCase) {
      console.log(`- ${tc.testCase.id}`);

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
            console.error(`Failed to fetch work item: ${tc.testCase.url}, status: ${response.status}`);
            continue;
          }
          // Define a minimal type for workItem
          type WorkItem = {
            fields?: {
              [key: string]: any;
            };
          };
          const workItem = (await response.json()) as WorkItem;
          const tags = workItem.fields?.["System.Tags"];
          if (tags && tags.includes("VSCUSE")) {
            const steps = workItem.fields?.["Microsoft.VSTS.TCM.Steps"];
            if (typeof steps === "string") {
              console.log(`TestCase ${tc.testCase.id} Steps:`);
              console.log(steps);
              const stepBlocks = steps.match(/<step[\s\S]*?<\/step>/gi) || [];
              console.log(`Found ${stepBlocks.length} step blocks.`);
              stepBlocks.forEach((stepBlock, idx) => {
                const paramMatch = stepBlock.match(/<parameterizedString[^>]*>([\s\S]*?)<\/parameterizedString>/i);
                if (paramMatch && paramMatch[1]) {
                  const html = paramMatch[1].replace(/&lt;/g, "<").replace(/&gt;/g, ">");
                  const pMatch = html.match(/<p>([\s\S]*?)<\/p>/i) || html.match(/<P>([\s\S]*?)<\/P>/i);
                  if (pMatch && pMatch[1]) {
                    const text = pMatch[1].replace(/<[^>]+>/g, "").trim();
                    console.log(`Step ${idx + 1}: ${text}`);
                  } else {
                    console.log(`Step ${idx + 1}: [No <p> found in parameterizedString]`);
                  }
                } else {
                  const match = stepBlock.match(/<p>([\s\S]*?)<\/p>/i) || stepBlock.match(/<P>([\s\S]*?)<\/P>/i);
                  if (match && match[1]) {
                    const text = match[1].replace(/<[^>]+>/g, "").trim();
                    console.log(`Step ${idx + 1}: ${text}`);
                  } else {
                    console.log(`Step ${idx + 1}: [No parameterizedString or <p> found]`);
                  }
                }
              });
            }
            else {
              console.log(`The type is: ${typeof steps}`);
            }
          }
        } catch (err) {
          console.error(`Error fetching work item for test case ${tc.testCase.id}:`, err);
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
