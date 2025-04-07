// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { ListAPIResult, ParseOptions, ProjectType, SpecParser } from "@microsoft/m365-spec-parser";
import { Platform } from "@microsoft/teamsfx-api";

const daProjectConfig: ParseOptions = {
  projectType: ProjectType.Copilot,
  isGptPlugin: true,
  allowMultipleParameters: true,
  allowMissingId: true,
  allowSwagger: true,
  allowAPIKeyAuth: true,
  allowBearerTokenAuth: true,
  allowOauth2: true,
  allowMethods: ["get", "post", "put", "delete", "patch", "head", "connect", "options", "trace"],
  allowResponseSemantics: true,
};

export async function listAPIInfo(specPath: string, platform?: string): Promise<ListAPIResult> {
  const options: ParseOptions = {
    ...daProjectConfig,
    allowAPIKeyAuth: platform !== Platform.VS,
    allowBearerTokenAuth: platform !== Platform.VS,
    allowOauth2: platform !== Platform.VS,
  };

  const parser = new SpecParser(specPath, options);

  return await parser.list();
}
