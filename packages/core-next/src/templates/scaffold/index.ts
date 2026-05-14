// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

export { TemplateInfo, ScaffoldContext, TemplateConfig, convertToLangKey } from "./types";

export { renderTemplateFileName, renderTemplateFileData } from "./render";
export { resolveTemplateUrl, fetchZip, unzipWithTransform, loadLocalFallback } from "./download";
export { scaffoldTemplates } from "./scaffolder";
export { getTemplateReplaceMap } from "./replaceMap";
