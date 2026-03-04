// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import {
  CLICommand,
  CLICommandOption,
  CLIContext,
  err,
  ok,
  Platform,
} from "@microsoft/teamsfx-api";
import {
  CliQuestionName,
  CreateProjectInputs,
  CreateProjectOptions,
  featureFlagManager,
  FeatureFlags,
  isTdpTemplate,
} from "@microsoft/teamsfx-core";
import chalk from "chalk";
import { assign } from "lodash";
import * as path from "path";
import * as uuid from "uuid";
import { getFxCore } from "../../activate";
import { logger } from "../../commonlib/logger";
import { commands } from "../../resource";
import { TelemetryEvent, TelemetryProperty } from "../../telemetry/cliTelemetryEvents";
import { createSampleCommand } from "./createSample";
import { listAllTemplates } from "./listTemplates";

function adjustOptions(options: CLICommandOption[]) {
  for (const option of options) {
    if (option.type === "string" && option.name === CliQuestionName.Capability) {
      // use dynamic options for capability question
      option.choices = listAllTemplates().flatMap((o) => (o.alias ? [o.alias, o.name] : [o.name]));
      break;
    }
  }

  return options;
}

export function getCreateCommand(): CLICommand {
  return {
    name: "new",
    description: commands.create.description,
    options: [...adjustOptions(CreateProjectOptions)],
    examples: [
      {
        command: `${process.env.TEAMSFX_CLI_BIN_NAME} new -c declarative-agent -n myagent -i false`,
        description: "Create a new declarative agent",
      },
      {
        command: `${process.env.TEAMSFX_CLI_BIN_NAME} new -c basic-custom-engine-agent -l typescript -n mycea -i false`,
        description: "Create a new basic custom engine agent",
      },
    ],
    commands: [createSampleCommand],
    telemetry: {
      event: TelemetryEvent.CreateProject,
    },
    handler: async (ctx: CLIContext) => {
      const inputs = ctx.optionValues as CreateProjectInputs;
      inputs.projectId = inputs.projectId ?? uuid.v4();
      const core = getFxCore();
      if (inputs.nonInteractive) {
        if (featureFlagManager.getBooleanValue(FeatureFlags.CLIDotNet)) {
          // this feature is used in e2e test to scaffold VS project in non-interactive mode
          inputs.platform = Platform.VS;
          inputs["template-name"] = inputs.capabilities;
          inputs["programming-language"] = "csharp";
        } else {
          // for non-interactive mode, we need to preset project-type from capability to make sure the question model works
          const capability = inputs.capabilities as string;
          inputs["template-name"] = capability;
          const templates = listAllTemplates();
          const matched = templates.find((t) => t.name === capability || t.alias === capability);
          if (matched) {
            inputs["template-name"] = matched.name;
            if (inputs["programming-language"] === undefined) {
              // preset programming language if not specified
              inputs["programming-language"] = matched.language as any;
            }
          }
        }
      }
      const isTdp = isTdpTemplate(inputs);
      const res = isTdp
        ? await core.createProjectFromTdp(inputs)
        : await core.createProject(inputs);
      assign(ctx.telemetryProperties, {
        [TelemetryProperty.NewProjectId]: inputs.projectId,
        [TelemetryProperty.IsCreatingM365]: inputs.isM365 + "",
      });
      if (res.isErr()) {
        return err(res.error);
      }
      logger.info(`Project created at: ${chalk.cyan(path.resolve(res.value.projectPath))}`);
      return ok(undefined);
    },
  };
}
