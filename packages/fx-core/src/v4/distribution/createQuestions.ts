// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { FxError, SystemError } from "@microsoft/teamsfx-api";
import AdmZip from "adm-zip";
import { Result, err, ok } from "neverthrow";
import { QuestionSpec } from "../collectInputs/collectInputs";
import { DeclarativeLocator } from "../model/dataModel";

/** Load one create template's `questions.json`. See collect-create-inputs spec. */

const SOURCE = "Scaffold";

/** The native question kinds `questions.json` may declare (collect-inputs `QuestionType`). */
const QUESTION_TYPES: ReadonlySet<string> = new Set([
  "singleSelect",
  "multiSelect",
  "text",
  "confirm",
  "singleFile",
  "folder",
  "singleFileOrText",
]);

/** The `v4/<kind>/<templateId>/questions.json` entry this locator resolves to. */
function questionsEntry(locator: DeclarativeLocator): string {
  return `v4/${locator.kind}/${locator.templateId}/questions.json`;
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

/** Predicate-narrow a parsed `questions` array without an unchecked cast. */
function isQuestionSpecArray(value: unknown): value is QuestionSpec[] {
  return (
    Array.isArray(value) &&
    value.every(
      (item) =>
        isRecord(item) &&
        typeof item.name === "string" &&
        typeof item.type === "string" &&
        QUESTION_TYPES.has(item.type)
    )
  );
}

export function openCreateQuestions(
  bytes: Buffer,
  locator: DeclarativeLocator
): Result<QuestionSpec[], FxError> {
  let zip: AdmZip;
  try {
    zip = new AdmZip(bytes);
  } catch {
    return err(
      new SystemError({
        source: SOURCE,
        name: "TemplatePackageCorrupt",
        message: "The resolved template package is not a valid archive.",
      })
    );
  }

  const entryPath = questionsEntry(locator);
  let questionsRaw: string | undefined;
  for (const entry of zip.getEntries()) {
    if (entry.isDirectory) {
      continue;
    }
    const name = entry.entryName.replace(/\\/g, "/");
    if (name === entryPath) {
      questionsRaw = entry.getData().toString("utf8");
      break;
    }
  }

  if (questionsRaw === undefined) {
    return err(
      new SystemError({
        source: SOURCE,
        name: "PackageFileMissing",
        message: `The template package is missing "${entryPath}".`,
      })
    );
  }

  let parsed: unknown;
  try {
    parsed = JSON.parse(questionsRaw);
  } catch {
    return err(
      new SystemError({
        source: SOURCE,
        name: "PackageFileInvalid",
        message: `The template package file "${entryPath}" is not valid JSON.`,
      })
    );
  }

  if (!isRecord(parsed) || !isQuestionSpecArray(parsed.questions)) {
    return err(
      new SystemError({
        source: SOURCE,
        name: "PackageFileInvalid",
        message: `The template package file "${entryPath}" must be an object with a "questions" array.`,
      })
    );
  }

  return ok(parsed.questions);
}
