// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { SystemError } from "@microsoft/teamsfx-api";
import type { FxError } from "@microsoft/teamsfx-api";
import type { Result } from "neverthrow";
import type { ReplaceMapEntry } from "../renderContext/buildRenderContext";
import type { TemplateFileEntry } from "../model/dataModel";
import type { Pipeline } from "../pipeline/runScaffoldPipeline";
import * as parser from "../validation/packageParse";
import type { PreparedTemplate } from "../validation/packageParse";

export { PACKAGE_PARSE_ERROR, parseDeclaredKeys } from "../validation/packageParse";
export type { PreparedTemplate } from "../validation/packageParse";

function systemError(name: string, message: string): SystemError {
  return new SystemError({ source: "Scaffold", name, message });
}

export function prepareTemplate(raw: {
  descriptor: unknown;
  pipeline: unknown;
  content: TemplateFileEntry[];
}): Result<PreparedTemplate, FxError> {
  return parser.prepareTemplate(raw, systemError);
}

/** Extract `descriptor.replaceMap` as a typed entry list (absent ⇒ empty). */
export function parseReplaceMap(descriptor: unknown): Result<ReplaceMapEntry[], FxError> {
  return parser.parseReplaceMap(descriptor, systemError);
}

/** Extract a parsed `pipeline.json` as a typed `Pipeline`. */
export function parsePipeline(raw: unknown): Result<Pipeline, FxError> {
  return parser.parsePipeline(raw, systemError);
}
