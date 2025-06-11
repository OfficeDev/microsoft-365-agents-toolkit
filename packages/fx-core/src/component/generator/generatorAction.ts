// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import AdmZip from "adm-zip";
import fs from "fs-extra";
import path from "path";

import { LogProvider, Platform } from "@microsoft/teamsfx-api";

import { getTemplatesFolder } from "../../folder";
import { MissKeyError, SampleNotFoundError, TemplateNotFoundError } from "./error";
import {
  downloadDirectory,
  fetchZipFromUrl,
  getSampleInfoFromName,
  unzip,
  getTemplateUrl,
  getTemplateLatestVersion,
} from "./utils";
import { SampleUrlInfo } from "../../common/samples";

export interface GeneratorContext {
  name: string;
  language?: string;
  platform?: Platform;
  destination: string;
  logProvider: LogProvider;
  tryLimits?: number;
  timeoutInMs?: number;
  sampleInfo?: SampleUrlInfo;
  fallback?: boolean;
  outputs?: string[];

  filterFn?: (name: string) => boolean;
  fileNameReplaceFn?: (name: string, data: Buffer) => string;
  fileDataReplaceFn?: (name: string, data: Buffer) => Buffer | string;

  onActionStart?: (action: GeneratorAction, context: GeneratorContext) => Promise<void>;
  onActionEnd?: (action: GeneratorAction, context: GeneratorContext) => Promise<void>;
  onActionError: (
    action: GeneratorAction,
    context: GeneratorContext,
    error: Error
  ) => Promise<void>;
}

export interface GeneratorAction {
  name: string;
  run: (context: GeneratorContext) => Promise<void>;
}

export enum GeneratorActionName {
  ScaffoldRemoteTemplate = "ScaffoldRemoteTemplate",
  ScaffoldLocalTemplate = "ScaffoldLocalTemplate",
  FetchSampleInfo = "FetchSampleInfo",
  DownloadDirectory = "DownloadDirectory",
}

export const ScaffoldRemoteTemplateAction: GeneratorAction = {
  name: GeneratorActionName.ScaffoldRemoteTemplate,
  run: async (context: GeneratorContext) => {
    if (!context.language) {
      throw new MissKeyError("language");
    }

    const templateUrl = await getTemplateUrl(
      context.language,
      getTemplateLatestVersion,
      context.platform
    );
    if (!templateUrl) {
      return;
    }

    const zip = await fetchZipFromUrl(templateUrl, context.tryLimits, context.timeoutInMs);
    context.outputs = await unzip(
      zip,
      context.destination,
      context.fileNameReplaceFn,
      context.fileDataReplaceFn,
      context.filterFn
    );
  },
};

export const ScaffoldLocalTemplateAction: GeneratorAction = {
  name: GeneratorActionName.ScaffoldLocalTemplate,
  run: async (context: GeneratorContext) => {
    if (!context.language) {
      throw new MissKeyError("language");
    }

    if (context.outputs?.length) {
      return;
    }
    context.logProvider.debug(`Fetching zip from local: ${JSON.stringify(context)}`);
    const fallbackPath = path.join(getTemplatesFolder(), "fallback");
    const fileName = `${context.language}.zip`;
    const zipPath: string = path.join(fallbackPath, fileName);

    const data: Buffer = await fs.readFile(zipPath);
    const zip = new AdmZip(data);
    context.outputs = await unzip(
      zip,
      context.destination,
      context.fileNameReplaceFn,
      context.fileDataReplaceFn,
      context.filterFn
    );

    if (!context.outputs?.length) {
      throw new TemplateNotFoundError(context.name);
    }
  },
};

export const fetchSampleInfoAction: GeneratorAction = {
  name: GeneratorActionName.FetchSampleInfo,
  run: async (context: GeneratorContext) => {
    const sample = await getSampleInfoFromName(context.name);
    context.sampleInfo = sample.downloadUrlInfo;
  },
};

export const downloadDirectoryAction: GeneratorAction = {
  name: GeneratorActionName.DownloadDirectory,
  run: async (context: GeneratorContext) => {
    context.logProvider.debug(`Downloading sample by directory: ${JSON.stringify(context)}`);
    if (!context.sampleInfo) {
      throw new MissKeyError("sampleInfo");
    }

    context.outputs = await downloadDirectory(context.sampleInfo, context.destination);
    if (!context.outputs?.length) {
      throw new SampleNotFoundError(context.name);
    }
  },
};

export const TemplateActionSeq: GeneratorAction[] = [
  ScaffoldRemoteTemplateAction,
  ScaffoldLocalTemplateAction,
];

export const SampleActionSeq: GeneratorAction[] = [fetchSampleInfoAction, downloadDirectoryAction];
