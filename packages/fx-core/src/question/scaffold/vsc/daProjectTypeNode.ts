// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { ErrorType, ProjectType, SpecParser } from "@microsoft/m365-spec-parser";
import { Inputs, IQTreeNode, SingleFileQuestion, UserError } from "@microsoft/teamsfx-api";
import * as os from "os";
import * as path from "path";
import { SpecParserSource } from "../../../common/constants";
import { Correlator } from "../../../common/correlator";
import { featureFlagManager, FeatureFlags } from "../../../common/featureFlags";
import { getLocalizedString } from "../../../common/localizeUtils";
import {
  ApiSpecTelemetryPropertis,
  getQuestionValidationErrorEventName,
  sendTelemetryErrorEvent,
} from "../../../common/telemetry";
import { pluginManifestUtils } from "../../../component/driver/teamsApp/utils/PluginManifestUtils";
import { getParserOptions } from "../../../component/generator/apiSpec/helper";
import { validateSourcePluginManifest } from "../../../component/generator/copilotExtension/helper";
import { CoreSource, FileNotSupportError } from "../../../error/common";
import { QuestionNames } from "../../constants";
import {
  ApiAuthOptions,
  ApiPluginStartOptions,
  DACapabilityOptions,
  setTemplateName,
} from "./CapabilityOptions";
import { ProjectTypeOptions } from "./ProjectTypeOptions";
import { apiSpecNode } from "./m365ProjectTypeNode";

export function pluginManifestQuestion(): SingleFileQuestion {
  const correlationId = Correlator.getId();
  return {
    type: "singleFile",
    name: QuestionNames.PluginManifestFilePath,
    title: getLocalizedString("core.createProjectQuestion.addExistingPlugin.pluginManifest.title"),
    cliDescription: "Plugin manifest path.",
    placeholder: getLocalizedString(
      "core.createProjectQuestion.addExistingPlugin.pluginManifest.placeholder"
    ),
    filters: {
      files: ["json"],
    },
    defaultFolder: os.homedir(),
    validation: {
      validFunc: async (input: string) => {
        const manifestRes = await pluginManifestUtils.readPluginManifestFile(input.trim());
        if (manifestRes.isErr()) {
          sendTelemetryErrorEvent(
            CoreSource,
            getQuestionValidationErrorEventName(QuestionNames.PluginManifestFilePath),
            manifestRes.error,
            {
              "correlation-id": correlationId,
            }
          );
          return (manifestRes.error as UserError).displayMessage;
        } else {
          const manifest = manifestRes.value;

          const checkRes = validateSourcePluginManifest(
            manifest,
            QuestionNames.PluginManifestFilePath
          );
          if (checkRes.isErr()) {
            sendTelemetryErrorEvent(
              CoreSource,
              getQuestionValidationErrorEventName(QuestionNames.PluginManifestFilePath),
              checkRes.error,
              {
                "correlation-id": correlationId,
              }
            );
            return checkRes.error.displayMessage;
          }
        }
      },
    },
  };
}

export function pluginApiSpecQuestion(): SingleFileQuestion {
  const correlationId = Correlator.getId();
  return {
    type: "singleFile",
    name: QuestionNames.PluginOpenApiSpecFilePath,
    title: getLocalizedString("core.createProjectQuestion.addExistingPlugin.apiSpec.title"),
    placeholder: getLocalizedString(
      "core.createProjectQuestion.addExistingPlugin.openApiSpec.placeholder"
    ),
    cliDescription: "OpenAPI description document used for your API plugin.",
    filters: {
      files: ["json", "yml", "yaml"],
    },
    defaultFolder: (inputs: Inputs) =>
      path.dirname(inputs[QuestionNames.PluginManifestFilePath] as string),
    validation: {
      validFunc: async (input: string, inputs?: Inputs) => {
        if (!inputs) {
          throw new Error("inputs is undefined"); // should never happen
        }
        const filePath = input.trim();

        const ext = path.extname(filePath).toLowerCase();
        if (![".json", ".yml", ".yaml"].includes(ext)) {
          const error = new FileNotSupportError(CoreSource, ["json", "yml", "yaml"].join(", "));
          sendTelemetryErrorEvent(
            CoreSource,
            getQuestionValidationErrorEventName(QuestionNames.PluginOpenApiSpecFilePath),
            error,
            {
              "correlation-id": correlationId,
            }
          );
          return error.displayMessage;
        }

        const specParser = new SpecParser(filePath, getParserOptions(ProjectType.Copilot));
        const validationRes = await specParser.validate();
        const invalidSpecError = validationRes.errors.find(
          (o) => o.type === ErrorType.SpecNotValid
        );

        if (invalidSpecError) {
          const error = new UserError(
            SpecParserSource,
            ApiSpecTelemetryPropertis.InvalidApiSpec,
            invalidSpecError.content,
            invalidSpecError.content
          );
          sendTelemetryErrorEvent(
            CoreSource,
            getQuestionValidationErrorEventName(QuestionNames.PluginOpenApiSpecFilePath),
            error,
            {
              "correlation-id": correlationId,
              [ApiSpecTelemetryPropertis.SpecNotValidDetails]: invalidSpecError.content,
            }
          );
        }
        return invalidSpecError?.content;
      },
    },
  };
}

export function declarativeAgentProjectTypeNode(
  parentValue = ProjectTypeOptions.copilotAgentOptionId
): IQTreeNode {
  return {
    // project-type = Declarative Agent
    condition: { equals: parentValue },
    data: {
      name: QuestionNames.WithPlugin,
      title: getLocalizedString("core.createProjectQuestion.declarativeCopilot.title"),
      cliDescription: "Whether to add API plugin for your declarative Copilot.",
      type: "singleSelect",
      staticOptions: [DACapabilityOptions.noPlugin(), DACapabilityOptions.withPlugin()],
      placeholder: getLocalizedString("core.createProjectQuestion.declarativeCopilot.placeholder"),
      onDidSelection: setTemplateName,
    },
    children: [
      {
        condition: { equals: DACapabilityOptions.withPlugin().id },
        data: {
          type: "singleSelect",
          name: QuestionNames.ApiPluginType,
          title: getLocalizedString("core.createProjectQuestion.createApiPlugin.title"),
          cliDescription: "API plugin type.",
          placeholder: getLocalizedString("core.createProjectQuestion.addApiPlugin.placeholder"),
          staticOptions: [
            ApiPluginStartOptions.newApi(),
            ApiPluginStartOptions.apiSpec(),
            ApiPluginStartOptions.existingPlugin(),
          ],
          default: ApiPluginStartOptions.newApi().id,
          onDidSelection: setTemplateName,
        },
        children: [
          {
            condition: { equals: ApiPluginStartOptions.newApi().id },
            data: {
              type: "singleSelect",
              name: QuestionNames.ApiAuth,
              title: getLocalizedString("core.createProjectQuestion.apiMessageExtensionAuth.title"),
              cliDescription: "The authentication type for the API.",
              placeholder: getLocalizedString(
                "core.createProjectQuestion.apiMessageExtensionAuth.placeholder"
              ),
              staticOptions: [
                ApiAuthOptions.none(),
                ApiAuthOptions.apiKey(),
                ApiAuthOptions.microsoftEntra(),
                ApiAuthOptions.oauth(),
              ],
              default: ApiAuthOptions.none().id,
              onDidSelection: setTemplateName,
            },
          },
          apiSpecNode(
            (inputs: Inputs) =>
              inputs[QuestionNames.ApiPluginType] === ApiPluginStartOptions.apiSpec().id &&
              !featureFlagManager.getBooleanValue(FeatureFlags.KiotaIntegration)
          ),
          {
            condition: { equals: ApiPluginStartOptions.existingPlugin().id },
            data: { type: "group", name: QuestionNames.ImportPlugin },
            children: [
              {
                data: pluginManifestQuestion(),
              },
              {
                data: pluginApiSpecQuestion(),
              },
            ],
          },
        ],
      },
    ],
  };
}
