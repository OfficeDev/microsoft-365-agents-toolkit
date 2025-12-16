// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Yuqi Zhou <yuqzho@microsoft.com>
 */

import * as chai from "chai";
import * as fs from "fs-extra";
import { describe } from "mocha";
import * as path from "path";

import { it } from "@microsoft/extra-shot-mocha";

import { CliHelper } from "../../commonlib/cliHelper";
import {
  cleanUpLocalProject,
  getTestFolder,
  getUniqueAppName,
  readContextMultiEnvV3,
} from "../commonUtils";
import {
  deleteAadAppByClientId,
  deleteBot,
  deleteTeamsApp,
  getAadAppByClientId,
  getBot,
  getTeamsApp,
} from "../debug/utility";

describe("Debug V3 custom-copilot-rag-ai-search TypeScript template", () => {
  const testFolder = getTestFolder();
  const appName = getUniqueAppName();
  const projectPath = path.resolve(testFolder, appName);

  afterEach(async function () {
    const context = await readContextMultiEnvV3(projectPath, "local");

    // clean up
    if (context?.TEAMS_APP_ID) {
      await deleteTeamsApp(context.TEAMS_APP_ID);
    }
    if (context?.BOT_ID) {
      await deleteBot(context.BOT_ID);
      await deleteAadAppByClientId(context.BOT_ID);
    }
    await cleanUpLocalProject(projectPath);
  });

  it(
    "Azure OpenAI happy path: provision and deploy",
    { testPlanCaseId: 27569074, author: "yuqzho@microsoft.com" },
    async function () {
      const myRecordAzOpenAI: Record<string, string> = {};
      myRecordAzOpenAI["programming-language"] = "typescript ";
      myRecordAzOpenAI["custom-copilot-rag"] =
        "custom-copilot-rag-azureAISearch";
      myRecordAzOpenAI["llm-service"] = "llm-service-azure-openai";
      myRecordAzOpenAI["azure-openai-key"] = "fake";
      myRecordAzOpenAI["azure-openai-deployment-name"] = "fake";
      myRecordAzOpenAI["azure-openai-endpoint"] = "https://test.com";
      const options = Object.entries(myRecordAzOpenAI)
        .map(([key, value]) => "--" + key + " " + value)
        .join(" ");
      await CliHelper.createProjectWithCapability(
        appName,
        testFolder,
        "custom-copilot-rag-azure-ai-search",
        undefined,
        options
      );
      console.log(`[Successfully] scaffold to ${projectPath}`);

      // add extra envs
      const userFile = path.resolve(projectPath, "env", `.env.local.user`);
      const AZURE_OPENAI_EMBEDDING_DEPLOYMENT =
        "AZURE_OPENAI_EMBEDDING_DEPLOYMENT=fake";
      const SECRET_AZURE_SEARCH_KEY = "SECRET_AZURE_SEARCH_KEY=fake";
      const AZURE_SEARCH_ENDPOINT = "AZURE_SEARCH_ENDPOINT=https://test.com";
      const KEY =
        "\n" +
        AZURE_OPENAI_EMBEDDING_DEPLOYMENT +
        "\n" +
        SECRET_AZURE_SEARCH_KEY +
        "\n" +
        AZURE_SEARCH_ENDPOINT;
      fs.appendFileSync(userFile, KEY);
      console.log(`add key ${KEY} to .env.local.user file`);

      /// provision
      await CliHelper.provisionProject(projectPath, "", "local", {
        ...process.env,
        BOT_DOMAIN: "test.ngrok.io",
        BOT_ENDPOINT: "https://test.ngrok.io",
      });
      console.log(`[Successfully] provision for ${projectPath}`);

      let context = await readContextMultiEnvV3(projectPath, "local");
      chai.assert.isDefined(context);

      // validate bot
      chai.assert.isDefined(context.BOT_ID);
      chai.assert.isNotEmpty(context.BOT_ID);
      const aadApp = await getAadAppByClientId(context.BOT_ID);
      chai.assert.isDefined(aadApp);
      chai.assert.equal(aadApp?.appId, context.BOT_ID);
      const bot = await getBot(context.BOT_ID);
      chai.assert.equal(bot?.botId, context.BOT_ID);
      chai.assert.equal(
        bot?.messagingEndpoint,
        "https://test.ngrok.io/api/messages"
      );
      chai.assert.deepEqual(bot?.configuredChannels, ["msteams"]);

      // validate teams app
      chai.assert.isDefined(context.TEAMS_APP_ID);
      const teamsApp = await getTeamsApp(context.TEAMS_APP_ID);
      chai.assert.equal(teamsApp?.teamsAppId, context.TEAMS_APP_ID);

      // deploy
      await CliHelper.deployAll(projectPath, "", "local");
      console.log(`[Successfully] deploy for ${projectPath}`);

      context = await readContextMultiEnvV3(projectPath, "local");
      chai.assert.isDefined(context);
    }
  );

  it(
    "OpenAI happy path: provision and deploy",
    { testPlanCaseId: 28970306, author: "yuqzho@microsoft.com" },
    async function () {
      // create
      const myRecordAzOpenAI: Record<string, string> = {};
      myRecordAzOpenAI["programming-language"] = "typescript ";
      myRecordAzOpenAI["custom-copilot-rag"] =
        "custom-copilot-rag-azureAISearch";
      myRecordAzOpenAI["llm-service"] = "llm-service-openai";
      myRecordAzOpenAI["openai-key"] = "fake";
      const options = Object.entries(myRecordAzOpenAI)
        .map(([key, value]) => "--" + key + " " + value)
        .join(" ");
      await CliHelper.createProjectWithCapability(
        appName,
        testFolder,
        "custom-copilot-rag-azure-ai-search",
        undefined,
        options
      );
      console.log(`[Successfully] scaffold to ${projectPath}`);

      // add extra envs
      const userFile = path.resolve(projectPath, "env", `.env.local.user`);
      const AZURE_OPENAI_EMBEDDING_DEPLOYMENT =
        "AZURE_OPENAI_EMBEDDING_DEPLOYMENT=fake";
      const SECRET_AZURE_SEARCH_KEY = "SECRET_AZURE_SEARCH_KEY=fake";
      const AZURE_SEARCH_ENDPOINT = "AZURE_SEARCH_ENDPOINT=https://test.com";
      const KEY =
        "\n" +
        AZURE_OPENAI_EMBEDDING_DEPLOYMENT +
        "\n" +
        SECRET_AZURE_SEARCH_KEY +
        "\n" +
        AZURE_SEARCH_ENDPOINT;
      fs.appendFileSync(userFile, KEY);
      console.log(`add key ${KEY} to .env.local.user file`);

      // provision
      await CliHelper.provisionProject(projectPath, "", "local", {
        ...process.env,
        BOT_DOMAIN: "test.ngrok.io",
        BOT_ENDPOINT: "https://test.ngrok.io",
      });
      console.log(`[Successfully] provision for ${projectPath}`);

      let context = await readContextMultiEnvV3(projectPath, "local");
      chai.assert.isDefined(context);

      // validate bot
      chai.assert.isDefined(context.BOT_ID);
      chai.assert.isNotEmpty(context.BOT_ID);
      const aadApp = await getAadAppByClientId(context.BOT_ID);
      chai.assert.isDefined(aadApp);
      chai.assert.equal(aadApp?.appId, context.BOT_ID);
      const bot = await getBot(context.BOT_ID);
      chai.assert.equal(bot?.botId, context.BOT_ID);
      chai.assert.equal(
        bot?.messagingEndpoint,
        "https://test.ngrok.io/api/messages"
      );
      chai.assert.deepEqual(bot?.configuredChannels, ["msteams"]);

      // validate teams app
      chai.assert.isDefined(context.TEAMS_APP_ID);
      const teamsApp = await getTeamsApp(context.TEAMS_APP_ID);
      chai.assert.equal(teamsApp?.teamsAppId, context.TEAMS_APP_ID);

      // deploy
      await CliHelper.deployAll(projectPath, "", "local");
      console.log(`[Successfully] deploy for ${projectPath}`);

      context = await readContextMultiEnvV3(projectPath, "local");
      chai.assert.isDefined(context);
    }
  );

  it(
    "JavaScript Azure OpenAI happy path: provision and deploy",
    { testPlanCaseId: 27569090, author: "yuqzho@microsoft.com" },
    async function () {
      // create
      const myRecordAzOpenAI: Record<string, string> = {};
      myRecordAzOpenAI["programming-language"] = "javascript ";
      myRecordAzOpenAI["custom-copilot-rag"] =
        "custom-copilot-rag-azureAISearch";
      myRecordAzOpenAI["llm-service"] = "llm-service-azure-openai";
      myRecordAzOpenAI["azure-openai-key"] = "fake";
      myRecordAzOpenAI["azure-openai-deployment-name"] = "fake";
      myRecordAzOpenAI["azure-openai-endpoint"] = "https://test.com";
      const options = Object.entries(myRecordAzOpenAI)
        .map(([key, value]) => "--" + key + " " + value)
        .join(" ");
      await CliHelper.createProjectWithCapability(
        appName,
        testFolder,
        "custom-copilot-rag-azure-ai-search",
        undefined,
        options
      );
      console.log(`[Successfully] scaffold to ${projectPath}`);

      // add extra envs
      const userFile = path.resolve(projectPath, "env", `.env.local.user`);
      const AZURE_OPENAI_EMBEDDING_DEPLOYMENT =
        "AZURE_OPENAI_EMBEDDING_DEPLOYMENT=fake";
      const SECRET_AZURE_SEARCH_KEY = "SECRET_AZURE_SEARCH_KEY=fake";
      const AZURE_SEARCH_ENDPOINT = "AZURE_SEARCH_ENDPOINT=https://test.com";
      const KEY =
        "\n" +
        AZURE_OPENAI_EMBEDDING_DEPLOYMENT +
        "\n" +
        SECRET_AZURE_SEARCH_KEY +
        "\n" +
        AZURE_SEARCH_ENDPOINT;
      fs.appendFileSync(userFile, KEY);
      console.log(`add key ${KEY} to .env.local.user file`);

      /// provision
      await CliHelper.provisionProject(projectPath, "", "local", {
        ...process.env,
        BOT_DOMAIN: "test.ngrok.io",
        BOT_ENDPOINT: "https://test.ngrok.io",
      });
      console.log(`[Successfully] provision for ${projectPath}`);

      let context = await readContextMultiEnvV3(projectPath, "local");
      chai.assert.isDefined(context);

      // validate bot
      chai.assert.isDefined(context.BOT_ID);
      chai.assert.isNotEmpty(context.BOT_ID);
      const aadApp = await getAadAppByClientId(context.BOT_ID);
      chai.assert.isDefined(aadApp);
      chai.assert.equal(aadApp?.appId, context.BOT_ID);
      const bot = await getBot(context.BOT_ID);
      chai.assert.equal(bot?.botId, context.BOT_ID);
      chai.assert.equal(
        bot?.messagingEndpoint,
        "https://test.ngrok.io/api/messages"
      );
      chai.assert.deepEqual(bot?.configuredChannels, ["msteams"]);

      // validate teams app
      chai.assert.isDefined(context.TEAMS_APP_ID);
      const teamsApp = await getTeamsApp(context.TEAMS_APP_ID);
      chai.assert.equal(teamsApp?.teamsAppId, context.TEAMS_APP_ID);

      // deploy
      await CliHelper.deployAll(projectPath, "", "local");
      console.log(`[Successfully] deploy for ${projectPath}`);

      context = await readContextMultiEnvV3(projectPath, "local");
      chai.assert.isDefined(context);
    }
  );

  it(
    "JavaScript OpenAI happy path: provision and deploy",
    { testPlanCaseId: 28970334, author: "yuqzho@microsoft.com" },
    async function () {
      // create
      // create
      const myRecordAzOpenAI: Record<string, string> = {};
      myRecordAzOpenAI["programming-language"] = "javascript ";
      myRecordAzOpenAI["custom-copilot-rag"] =
        "custom-copilot-rag-azureAISearch";
      myRecordAzOpenAI["llm-service"] = "llm-service-openai";
      myRecordAzOpenAI["openai-key"] = "fake";
      const options = Object.entries(myRecordAzOpenAI)
        .map(([key, value]) => "--" + key + " " + value)
        .join(" ");
      await CliHelper.createProjectWithCapability(
        appName,
        testFolder,
        "custom-copilot-rag-azure-ai-search",
        undefined,
        options
      );
      console.log(`[Successfully] scaffold to ${projectPath}`);

      // add extra envs
      const userFile = path.resolve(projectPath, "env", `.env.local.user`);
      const AZURE_OPENAI_EMBEDDING_DEPLOYMENT =
        "AZURE_OPENAI_EMBEDDING_DEPLOYMENT=fake";
      const SECRET_AZURE_SEARCH_KEY = "SECRET_AZURE_SEARCH_KEY=fake";
      const AZURE_SEARCH_ENDPOINT = "AZURE_SEARCH_ENDPOINT=https://test.com";
      const KEY =
        "\n" +
        AZURE_OPENAI_EMBEDDING_DEPLOYMENT +
        "\n" +
        SECRET_AZURE_SEARCH_KEY +
        "\n" +
        AZURE_SEARCH_ENDPOINT;
      fs.appendFileSync(userFile, KEY);
      console.log(`add key ${KEY} to .env.local.user file`);

      // provision
      await CliHelper.provisionProject(projectPath, "", "local", {
        ...process.env,
        BOT_DOMAIN: "test.ngrok.io",
        BOT_ENDPOINT: "https://test.ngrok.io",
      });
      console.log(`[Successfully] provision for ${projectPath}`);

      let context = await readContextMultiEnvV3(projectPath, "local");
      chai.assert.isDefined(context);

      // validate bot
      chai.assert.isDefined(context.BOT_ID);
      chai.assert.isNotEmpty(context.BOT_ID);
      const aadApp = await getAadAppByClientId(context.BOT_ID);
      chai.assert.isDefined(aadApp);
      chai.assert.equal(aadApp?.appId, context.BOT_ID);
      const bot = await getBot(context.BOT_ID);
      chai.assert.equal(bot?.botId, context.BOT_ID);
      chai.assert.equal(
        bot?.messagingEndpoint,
        "https://test.ngrok.io/api/messages"
      );
      chai.assert.deepEqual(bot?.configuredChannels, ["msteams"]);

      // validate teams app
      chai.assert.isDefined(context.TEAMS_APP_ID);
      const teamsApp = await getTeamsApp(context.TEAMS_APP_ID);
      chai.assert.equal(teamsApp?.teamsAppId, context.TEAMS_APP_ID);

      // deploy
      await CliHelper.deployAll(projectPath, "", "local");
      console.log(`[Successfully] deploy for ${projectPath}`);

      context = await readContextMultiEnvV3(projectPath, "local");
      chai.assert.isDefined(context);
    }
  );
});
