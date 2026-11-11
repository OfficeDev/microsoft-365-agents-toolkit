// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { FxError, SystemError, UserError } from "@microsoft/teamsfx-api";
import { Result, err, ok } from "neverthrow";
import { capabilityDeclarations } from "../capabilities/declarations";
import { OptionItem, OptionsProvider } from "../collectInputs/collectInputs";
import { readDescriptorLanguages } from "../distribution/descriptorLanguages";
import { getLocalizedString } from "../../common/localizeUtils";

export const CREATE_LANGUAGES_PROVIDER = capabilityDeclarations.provider.createLanguages.id;

const LANGUAGE_LABELS: Record<string, string> = {
  javascript: "JavaScript",
  typescript: "TypeScript",
  csharp: "C#",
  python: "Python",
};

export interface CreateLanguageContext {
  descriptor?: unknown;
  surface?: string;
  flagReader?: (name: string) => boolean;
}

export function gateLanguagesBySurface(
  languages: string[],
  surface: string,
  flagReader: (name: string) => boolean
): string[] {
  const allowCsharp = surface !== "vscode" && flagReader("TEAMSFX_CLI_DOTNET");
  return allowCsharp ? languages : languages.filter((language) => language !== "csharp");
}

export function createLanguageOptionsProvider(
  context: CreateLanguageContext = {}
): OptionsProvider {
  return {
    fetch() {
      const data = readDescriptorLanguages(context.descriptor);
      if ("error" in data) {
        throw new SystemError({
          source: "Scaffold",
          name: "TemplateLanguageOptionsInvalid",
          message: getLocalizedString("core.v4.scaffold.languageOptionsInvalid"),
          error: new Error(data.error),
        });
      }
      const allowed = new Set(
        gateLanguagesBySurface(
          data.languages,
          context.surface ?? "vscode",
          context.flagReader ?? (() => false)
        )
      );
      return {
        options: data.options
          .filter((option) => allowed.has(option.id))
          .map((option) => ({ label: LANGUAGE_LABELS[option.id] ?? option.id, ...option })),
      };
    },
  };
}

export async function resolveLanguageOptions(
  provider: OptionsProvider
): Promise<Result<OptionItem[], FxError>> {
  try {
    return ok((await provider.fetch({})).options);
  } catch (error) {
    return err(
      error instanceof UserError || error instanceof SystemError
        ? error
        : new SystemError({
            source: "Scaffold",
            name: "InputProviderFailed",
            message: getLocalizedString("core.v4.scaffold.languageOptionsFailed"),
            error: error instanceof Error ? error : undefined,
          })
    );
  }
}
