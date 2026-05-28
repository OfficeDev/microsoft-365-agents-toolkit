// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import axios from "axios";
import fs from "fs-extra";
import { includes } from "lodash";
import path from "path";
import * as tmp from "tmp";

import {
  getSampleFileInfo,
  runWithLimitedConcurrency,
  sampleProvider,
  sendRequestWithRetry,
} from "@microsoft/teamsfx-core";
import {
  CancellationToken,
  ChatRequest,
  ChatResponseFileTree,
  ChatResponseStream,
  LanguageModelChatMessage,
  Uri,
} from "vscode";
import { getSampleMatchChatMessages, getTemplateMatchChatMessages } from "../../prompts";
import { IChatTelemetryData } from "../../types";
import {
  countMessagesTokens,
  getCopilotResponseAsString,
  getSampleDownloadUrlInfo,
} from "../../utils";
import teamsTemplateMetadata from "./templateMetadata.json";
import { ProjectMetadata } from "./types";

const TOKEN_LIMITS = 2700;
const SCORE_LIMIT = 0.6;

export async function matchProject(
  request: ChatRequest,
  token: CancellationToken,
  telemetryMetadata: IChatTelemetryData
): Promise<ProjectMetadata[]> {
  const allProjectMetadata = [...getTeamsTemplateMetadata(), ...(await getTeamsSampleMetadata())];
  const matchedProjects = [...(await matchTemplates(request, token, telemetryMetadata))];
  // using hard-coded "template" to narrow down search scope
  if (!request.prompt.includes("template")) {
    // also search in samples
    matchedProjects.push(...(await matchSamples(request, token, telemetryMetadata)));
    matchedProjects.sort((a, b) => b.score - a.score);
  } else {
    matchedProjects.sort((a, b) => b.score - a.score);
    matchedProjects.splice(2);
  }
  const result: ProjectMetadata[] = [];
  const matchedProjectIds = new Set<string>();
  for (const { id, score } of matchedProjects) {
    if (score < SCORE_LIMIT) {
      break;
    }
    const matchedProject = allProjectMetadata.find((config) => config.id === id);
    if (matchedProject && !matchedProjectIds.has(matchedProject.id)) {
      result.push(matchedProject);
      matchedProjectIds.add(matchedProject.id);
    }
  }
  return result;
}

async function matchTemplates(
  request: ChatRequest,
  token: CancellationToken,
  telemetryMetadata: IChatTelemetryData
): Promise<Array<{ id: string; score: number }>> {
  const templateExamples = [
    {
      user: "an app shown in sharepoint",
      app: '{"app":[{"id":"tab-spfx","score":1.0}]}',
    },
    {
      user: "an office addin",
      app: '{"app":[{"id":"outlook-addin-type","score":1.0}]}',
    },
    {
      user: "a bot app",
      app: '{"app":[{"id":"bot","score":1.0},{"id":"notification","score":0.7},{"id":"command-bot","score": 0.8},{"id":"workflow-bot","score":0.8}]}',
    },
    {
      user: "a Word addin",
      app: '{"app": []}',
    },
  ];
  const templateMetadata = getTeamsTemplateMetadata();
  const matchedTemplates = await sendCopilotMatchRequest(
    getTemplateMatchChatMessages(templateMetadata, templateExamples, request.prompt),
    token,
    telemetryMetadata
  );
  return matchedTemplates;
}

async function matchSamples(
  request: ChatRequest,
  token: CancellationToken,
  telemetryMetadata: IChatTelemetryData
): Promise<Array<{ id: string; score: number }>> {
  const sampleMetadata = await getTeamsSampleMetadata();
  const sampleExamples = [
    {
      user: "an app that manages to-do list and works in Outlook",
      app: '{"app": [{"id": "todo-list-with-Azure-backend-M365", "score": 1.0}]}',
    },
    {
      user: "an app to send notification to a lot of users",
      app: '{"app": [{"id": "large-scale-notification", "score": 1.0}]}',
    },
    {
      user: "a calculator app",
      app: '{"app": []}',
    },
  ];
  const exampleIds = sampleExamples.map((example) => example.app);
  const sampleExampleMetadata = sampleMetadata.filter((config) => includes(exampleIds, config.id));
  const remainingSampleMetadata = sampleMetadata.filter(
    (config) => !includes(exampleIds, config.id)
  );
  let index = 0;
  let projectMetadata: ProjectMetadata[] = [...sampleExampleMetadata];
  const matchedSamples: Array<{ id: string; score: number }> = [];
  while (index < remainingSampleMetadata.length) {
    projectMetadata.push(remainingSampleMetadata[index]);
    index += 1;
    const messages = getSampleMatchChatMessages(projectMetadata, sampleExamples, request.prompt);
    const tokenNumber = countMessagesTokens(messages);
    if (tokenNumber > TOKEN_LIMITS) {
      matchedSamples.push(...(await sendCopilotMatchRequest(messages, token, telemetryMetadata)));
      projectMetadata = [...sampleExampleMetadata];
    }
  }
  if (projectMetadata.length > sampleExampleMetadata.length) {
    matchedSamples.push(
      ...(await sendCopilotMatchRequest(
        getSampleMatchChatMessages(projectMetadata, sampleExamples, request.prompt),
        token,
        telemetryMetadata
      ))
    );
  }
  return matchedSamples;
}

async function sendCopilotMatchRequest(
  messages: LanguageModelChatMessage[],
  token: CancellationToken,
  telemetryMetadata: IChatTelemetryData
) {
  telemetryMetadata.chatMessages.push(...messages);

  const response = await getCopilotResponseAsString("copilot-gpt-4", messages, token);

  if (response) {
    try {
      const responseJson = JSON.parse(response);
      if (responseJson && responseJson.app) {
        return responseJson.app as Array<{ id: string; score: number }>;
      }
    } catch (e) {}
  }
  return [];
}

export function getTeamsTemplateMetadata(): ProjectMetadata[] {
  return teamsTemplateMetadata.map((config) => {
    return {
      id: config.id,
      type: "template",
      platform: "Teams",
      name: config.name,
      description: config.description,
      data: {
        capabilities: config.id,
        "project-type": config["project-type"],
      },
    };
  });
}

export async function getTeamsSampleMetadata(): Promise<ProjectMetadata[]> {
  const sampleCollection = await sampleProvider.SampleCollection;
  const result: ProjectMetadata[] = [];
  for (const sample of sampleCollection.samples) {
    result.push({
      id: sample.id,
      type: "sample",
      platform: "Teams",
      name: sample.title,
      description: sample.fullDescription,
    });
  }
  return result;
}

export async function showFileTree(
  projectMetadata: ProjectMetadata,
  response: ChatResponseStream
): Promise<string> {
  const downloadUrlInfo = await getSampleDownloadUrlInfo(projectMetadata.id);
  const { samplePaths, fileUrlPrefix } = await getSampleFileInfo(downloadUrlInfo, 2);
  const tempFolder = tmp.dirSync({ unsafeCleanup: true }).name;
  const nodes = await buildFileTree(
    fileUrlPrefix,
    samplePaths,
    tempFolder,
    downloadUrlInfo.dir,
    2,
    20
  );
  response.filetree(nodes, Uri.file(path.join(tempFolder, downloadUrlInfo.dir)));
  return path.join(tempFolder, downloadUrlInfo.dir);
}

export async function buildFileTree(
  fileUrlPrefix: string,
  samplePaths: string[],
  dstPath: string,
  relativeFolderName: string,
  retryLimits: number,
  concurrencyLimits: number
): Promise<ChatResponseFileTree[]> {
  const root: ChatResponseFileTree = {
    name: relativeFolderName,
    children: [],
  };
  const downloadCallback = async (samplePath: string) => {
    const file = (await sendRequestWithRetry(async () => {
      return await axios.get(fileUrlPrefix + samplePath, {
        responseType: "arraybuffer",
      });
    }, retryLimits)) as unknown as any;
    const relativePath = path.relative(`${relativeFolderName}/`, samplePath);
    const filePath = path.join(dstPath, samplePath);
    fileTreeAdd(root, relativePath);
    await fs.ensureFile(filePath);
    await fs.writeFile(filePath, Buffer.from(file.data));
  };
  await runWithLimitedConcurrency(samplePaths, downloadCallback, concurrencyLimits);
  return root.children ?? [];
}

export function fileTreeAdd(root: ChatResponseFileTree, relativePath: string) {
  const filename = path.basename(relativePath);
  const folderName = path.dirname(relativePath);
  const segments = path.sep === "\\" ? folderName.split("\\") : folderName.split("/");
  let parent = root;
  for (let i = 0; i < segments.length; i++) {
    const segment = segments[i];
    if (segment === ".") {
      continue;
    }
    let child = parent.children?.find((child) => child.name === segment);
    if (!child) {
      child = {
        name: segment,
        children: [],
      };
      parent.children?.push(child);
    }
    parent = child;
  }
  parent.children?.push({
    name: filename,
  });
}
