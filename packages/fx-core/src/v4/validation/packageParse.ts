// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import type { FxError } from "@microsoft/teamsfx-api";
import { err, ok } from "neverthrow";
import type { Result } from "neverthrow";
import type { ConditionalExpression } from "../expression/evaluateExpression";
import type { ReplaceMapEntry } from "../renderContext/buildRenderContext";
import type { TemplateFileEntry } from "../model/dataModel";
import { readDescriptorLanguages } from "../distribution/descriptorLanguages";
import type {
  Pipeline,
  PipelineRender,
  PipelineStep,
  RenderFilter,
  StepParams,
} from "../pipeline/runScaffoldPipeline";

export type PackageParseErrorFactory = (name: string, message: string) => FxError;

export interface PreparedTemplate {
  descriptor: {
    id?: string;
    minEngineVersion?: string;
    languages: string[];
    replaceMap: ReplaceMapEntry[];
    declaredKeys: string[];
  };
  pipeline: Pipeline;
  content: TemplateFileEntry[];
}

export function prepareTemplate(
  raw: {
    descriptor: unknown;
    pipeline: unknown;
    content: TemplateFileEntry[];
  },
  errorFactory: PackageParseErrorFactory
): Result<PreparedTemplate, FxError> {
  const replaceMap = parseReplaceMap(raw.descriptor, errorFactory);
  if (replaceMap.isErr()) return err(replaceMap.error);
  const pipeline = parsePipeline(raw.pipeline, errorFactory);
  if (pipeline.isErr()) return err(pipeline.error);
  const languages = readDescriptorLanguages(raw.descriptor);
  if ("error" in languages) return err(errorFactory(PACKAGE_PARSE_ERROR, languages.error));
  const metadata = isRecord(raw.descriptor) ? raw.descriptor : {};
  return ok({
    descriptor: {
      id: stringField(metadata, "id"),
      minEngineVersion: stringField(metadata, "minEngineVersion"),
      languages: languages.languages,
      replaceMap: replaceMap.value,
      declaredKeys: parseDeclaredKeys(raw.descriptor),
    },
    pipeline: pipeline.value,
    content: raw.content,
  });
}

export const PACKAGE_PARSE_ERROR = "TemplatePackageParseError";

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function stringField(record: Record<string, unknown>, key: string): string | undefined {
  const value = record[key];
  return typeof value === "string" ? value : undefined;
}

function toReplaceMapEntry(item: unknown): ReplaceMapEntry | undefined {
  if (!isRecord(item)) {
    return undefined;
  }
  const varName = stringField(item, "var");
  if (varName === undefined) {
    return undefined;
  }
  const constVal = stringField(item, "const");
  if (constVal !== undefined) {
    return { var: varName, const: constVal };
  }
  const fromVal = stringField(item, "from");
  if (fromVal !== undefined) {
    return { var: varName, from: fromVal };
  }
  const exprVal = stringField(item, "expr");
  if (exprVal !== undefined) {
    return { var: varName, expr: exprVal };
  }
  const whenVal = stringField(item, "when");
  const valueVal = stringField(item, "value");
  if (whenVal !== undefined && valueVal !== undefined) {
    return { var: varName, when: whenVal, value: valueVal };
  }
  return undefined;
}

export function parseReplaceMap(
  descriptor: unknown,
  errorFactory: PackageParseErrorFactory
): Result<ReplaceMapEntry[], FxError> {
  if (!isRecord(descriptor)) {
    return err(errorFactory(PACKAGE_PARSE_ERROR, "descriptor.json must be a JSON object"));
  }
  const raw = descriptor.replaceMap;
  if (raw === undefined) {
    return ok([]);
  }
  if (!Array.isArray(raw)) {
    return err(errorFactory(PACKAGE_PARSE_ERROR, "descriptor.replaceMap must be an array"));
  }
  const entries: ReplaceMapEntry[] = [];
  for (const item of raw) {
    const entry = toReplaceMapEntry(item);
    if (entry === undefined) {
      return err(
        errorFactory(PACKAGE_PARSE_ERROR, `invalid replaceMap entry: ${JSON.stringify(item)}`)
      );
    }
    entries.push(entry);
  }
  return ok(entries);
}

export function parseDeclaredKeys(descriptor: unknown): string[] {
  if (!isRecord(descriptor)) {
    return [];
  }
  const optionsSchema = descriptor.optionsSchema;
  if (!isRecord(optionsSchema)) {
    return [];
  }
  const properties = optionsSchema.properties;
  if (!isRecord(properties)) {
    return [];
  }
  return Object.keys(properties);
}

function toStringArray(raw: unknown): string[] | undefined {
  if (!Array.isArray(raw)) {
    return undefined;
  }
  const out: string[] = [];
  for (const value of raw) {
    if (typeof value !== "string") {
      return undefined;
    }
    out.push(value);
  }
  return out;
}

function toStepParams(raw: unknown): StepParams | undefined {
  if (!isRecord(raw)) {
    return undefined;
  }
  const params: StepParams = {};
  for (const [key, value] of Object.entries(raw)) {
    if (typeof value === "string" || typeof value === "boolean") {
      params[key] = value;
      continue;
    }
    const stringArray = toStringArray(value);
    if (stringArray === undefined) {
      return undefined;
    }
    params[key] = stringArray;
  }
  return params;
}

function applyConditionalMetadata(
  item: Record<string, unknown>,
  target: ConditionalExpression
): boolean {
  const comment = stringField(item, "comment");
  if (comment !== undefined) {
    target.comment = comment;
  }
  const when = stringField(item, "when");
  if (when !== undefined) {
    target.when = when;
  } else if (item.when !== undefined) {
    return false;
  }
  return true;
}

function toPipelineStep(item: unknown): PipelineStep | undefined {
  if (!isRecord(item)) {
    return undefined;
  }
  const stepName = stringField(item, "step");
  if (stepName === undefined) {
    return undefined;
  }
  const result: PipelineStep = { step: stepName };
  if (!applyConditionalMetadata(item, result)) {
    return undefined;
  }
  if (item.with !== undefined) {
    const params = toStepParams(item.with);
    if (params === undefined) {
      return undefined;
    }
    result.with = params;
  }
  if (item.produces !== undefined) {
    const produces = toStringArray(item.produces);
    if (produces === undefined) {
      return undefined;
    }
    result.produces = produces;
  }
  return result;
}

function toRenderFilter(item: unknown): RenderFilter | undefined {
  if (!isRecord(item)) {
    return undefined;
  }
  const exclude = toStringArray(item.exclude);
  if (exclude === undefined) {
    return undefined;
  }
  const result: RenderFilter = { exclude };
  if (!applyConditionalMetadata(item, result)) {
    return undefined;
  }
  return result;
}

function toPipelineRender(raw: unknown): PipelineRender | undefined {
  if (!isRecord(raw)) {
    return undefined;
  }
  const result: PipelineRender = {};
  if (raw.filters !== undefined) {
    if (!Array.isArray(raw.filters)) {
      return undefined;
    }
    const filters: RenderFilter[] = [];
    for (const item of raw.filters) {
      const filter = toRenderFilter(item);
      if (filter === undefined) {
        return undefined;
      }
      filters.push(filter);
    }
    result.filters = filters;
  }
  return result;
}

export function parsePipeline(
  raw: unknown,
  errorFactory: PackageParseErrorFactory
): Result<Pipeline, FxError> {
  if (!isRecord(raw)) {
    return err(errorFactory(PACKAGE_PARSE_ERROR, "pipeline.json must be a JSON object"));
  }
  const name = stringField(raw, "pipeline");
  if (name === undefined) {
    return err(
      errorFactory(PACKAGE_PARSE_ERROR, "pipeline.pipeline (the orchestration name) is required")
    );
  }
  if (!Array.isArray(raw.steps)) {
    return err(errorFactory(PACKAGE_PARSE_ERROR, "pipeline.steps must be an array"));
  }
  const steps: PipelineStep[] = [];
  for (const item of raw.steps) {
    const step = toPipelineStep(item);
    if (step === undefined) {
      return err(
        errorFactory(PACKAGE_PARSE_ERROR, `invalid pipeline step: ${JSON.stringify(item)}`)
      );
    }
    steps.push(step);
  }
  const comment = stringField(raw, "comment");
  let render: PipelineRender | undefined;
  if (raw.render !== undefined) {
    render = toPipelineRender(raw.render);
    if (render === undefined) {
      return err(
        errorFactory(PACKAGE_PARSE_ERROR, `invalid pipeline render: ${JSON.stringify(raw.render)}`)
      );
    }
  }
  return ok({
    pipeline: name,
    ...(comment !== undefined ? { comment } : {}),
    ...(render !== undefined ? { render } : {}),
    steps,
  });
}
