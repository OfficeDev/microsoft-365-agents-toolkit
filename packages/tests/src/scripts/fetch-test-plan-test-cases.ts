import * as azdev from "azure-devops-node-api";
import { AzureCliCredential } from "@azure/identity";

async function run() {
  // Print environment variables for debugging
  console.error("Environment Variables:");
  console.error("AZURE_DEVOPS_ORG_URL:", process.env.AZURE_DEVOPS_ORG_URL);
  console.error("AZURE_DEVOPS_PROJECT:", process.env.AZURE_DEVOPS_PROJECT);
  console.error("AZURE_DEVOPS_TEST_PLAN_ID:", process.env.AZURE_DEVOPS_TEST_PLAN_ID);
  console.error("AZURE_DEVOPS_TEST_SUITE_ID:", process.env.AZURE_DEVOPS_TEST_SUITE_ID);
  const orgUrl = process.env.AZURE_DEVOPS_ORG_URL!;
  const project = process.env.AZURE_DEVOPS_PROJECT!;
  const planId = parseInt(process.env.AZURE_DEVOPS_TEST_PLAN_ID!);
  const suiteId = parseInt(process.env.AZURE_DEVOPS_TEST_SUITE_ID!);


  // Get Azure AD token using Azure CLI credential
  const credential = new AzureCliCredential();
  const token = await credential.getToken("https://app.vssps.visualstudio.com/.default");
  // Log the token for debugging (optional, be cautious with sensitive data)
  console.error("Azure AD Token GOT success, its length is:", token.token.length);
  // Initialize Azure DevOps API client

  // Create auth handler using the Azure AD token
  const authHandler = azdev.getBearerHandler(token.token);
  // log the auth handler for debugging
  console.error("Auth Handler created successfully");
  const connection = new azdev.WebApi(orgUrl, authHandler);
  console.error("Connection to Azure DevOps established successfully");
  const testApi = await connection.getTestApi();
  console.error("Test API client initialized successfully");

  const testCases = await testApi.getTestCases(project, planId, suiteId);
  console.error("Test cases fetched successfully");
  console.error(`Found ${testCases.length} test cases in plan ${planId}, suite ${suiteId}`);
  for (const tc of testCases) {
    console.error(`- ${tc.testCase.id}: ${tc.testCase.name}`);
  }
}

run().catch((err) => {
  console.error("Error:", err);
  process.exit(1);
});
