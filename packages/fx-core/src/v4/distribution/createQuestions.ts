// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { FxError, SystemError } from "@microsoft/teamsfx-api";
import AdmZip from "adm-zip";
import { Result, err, ok } from "neverthrow";
import { QuestionSpec } from "../collectInputs/collectInputs";
import { DeclarativeLocator } from "../model/dataModel";

/**
 * The `questions.json` load face of one create template — the Q2 sibling of
 * `openDeclarativePackage` (which reads only `descriptor.json` / `pipeline.json`
 * / `content/**`). Open the channel `templates.zip` bytes, read the single
 * `v4/<kind>/<templateId>/questions.json` entry, and project its
 * `{ questions: QuestionSpec[] }` envelope onto the native `QuestionSpec[]` the
 * surface-neutral driver (`collect-inputs`) walks.
 *
 * Spec: docs/03-specs/operations/scaffolding/collect-create-inputs.md (CCI-09).
 *
 * The shape extraction is predicate-narrowed (no `as`): a record envelope whose
 * `questions` is an array of `{ name, type }` records narrows to
 * `QuestionSpec[]`. Full per-field validation is the build-time
 * `validate-template-package`'s job — at runtime the authored, schema-valid
 * `questions.json` is trusted, the same untrusted-JSON discipline the other v4
 * boundary readers use.
 *
 * The zip-read faults — a non-archive byte string, a floor missing the entry, a
 * non-JSON / malformed entry — are `SystemError`s (packaging faults, not
 * user-fixable).
 *
 * v4-owned (INV-7): imports no v3 symbol; v3 may call `openCreateQuestions`, but
 * nothing here is tailored for v3.
 */

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

/**
 * Narrow a parsed `questions` array to `QuestionSpec[]`: every element is a
 * record carrying a string `name` and a known `type`. A user-defined type guard
 * (not an `as` cast); the remaining authored fields are trusted to the build
 * gate (`validate-template-package`).
 */
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
