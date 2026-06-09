// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { BaseComponentInnerError } from "../error/componentError";
import { errorSource } from "./constant";
import axios from "axios";
import { simplifyAxiosError } from "./utils";

export class CancelDownloading extends Error {}

export class SampleNotFoundError extends BaseComponentInnerError {
  constructor(templateName: string) {
    super(
      errorSource,
      "SystemError",
      "SampleNotFoundError",
      "error.generator.SampleNotFoundError",
      [templateName]
    );
  }
}

export class TemplateNotFoundError extends BaseComponentInnerError {
  constructor(templateName: string) {
    super(
      errorSource,
      "SystemError",
      "TemplateNotFoundError",
      "error.generator.TemplateNotFoundError",
      [templateName]
    );
  }
}

export class ScaffoldLocalTemplateError extends BaseComponentInnerError {
  constructor(error: Error) {
    super(
      errorSource,
      "SystemError",
      "ScaffoldLocalTemplateError",
      "error.generator.ScaffoldLocalTemplateError",
      undefined,
      undefined,
      undefined,
      undefined,
      error
    );
  }
}

export class UnzipError extends BaseComponentInnerError {
  constructor() {
    super(errorSource, "SystemError", "UnzipError", "error.generator.UnzipError", undefined, [
      "plugins.frontend.checkFsPermissionsTip",
    ]);
  }
}

export class TemplateOutputPathError extends BaseComponentInnerError {
  constructor(entryName: string) {
    super(
      errorSource,
      "SystemError",
      "TemplateOutputPathError",
      "error.generator.TemplateOutputPathError",
      [entryName]
    );
  }
}

export class DownloadSampleNetworkError extends BaseComponentInnerError {
  constructor(url: string, error: Error) {
    const innerError = axios.isAxiosError(error) ? simplifyAxiosError(error) : error;
    super(
      errorSource,
      "UserError",
      "DownloadSampleNetworkError",
      "error.generator.DownloadSampleNetworkError",
      [url],
      undefined,
      undefined,
      undefined,
      innerError
    );
  }
}
export class FetchSampleInfoError extends BaseComponentInnerError {
  constructor(error: Error) {
    const innerError = axios.isAxiosError(error) ? simplifyAxiosError(error) : error;
    super(
      errorSource,
      "UserError",
      "FetchSampleInfoError",
      "error.generator.FetchSampleInfoError",
      undefined,
      undefined,
      undefined,
      undefined,
      innerError
    );
  }
}

export class DownloadSampleApiLimitError extends BaseComponentInnerError {
  constructor(url: string, error: Error) {
    const innerError = axios.isAxiosError(error) ? simplifyAxiosError(error) : error;
    super(
      errorSource,
      "UserError",
      "DownloadSampleApiLimitError",
      "error.generator.DownloadSampleApiLimitError",
      [url],
      undefined,
      undefined,
      undefined,
      innerError
    );
  }
}

export class MissKeyError extends BaseComponentInnerError {
  constructor(keyName: string) {
    super(errorSource, "SystemError", "MissKeyError", "error.generator.MissKeyError", [keyName]);
  }
}
