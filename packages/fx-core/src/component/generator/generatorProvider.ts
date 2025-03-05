// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import { DeclarativeAgentGenerator } from "./declarativeAgent/generator";
import { DefaultTemplateGenerator } from "./defaultGenerator";
import { OfficeAddinGeneratorNew } from "./officeAddin/generator";
import { CustomEngineAgentWithExistingApiSpecGenerator } from "./openApiSpec/customEngineAgentGenerator";
import { DeclarativeAgentWithExistingApiSpecGenerator } from "./openApiSpec/declarativeAgentGenerator";
import { MessageExtensionWithExistingApiSpecGenerator } from "./openApiSpec/messageExtensionGenerator";
import { SsrTabGenerator } from "./other/ssrTabGenerator";
import { TdpGenerator } from "./other/tdpGenerator";
import { SPFxGeneratorImport, SPFxGeneratorNew } from "./spfx/spfxGenerator";

// When multiple generators are activated, only the top one will be executed.
export const Generators = [
  // TDP is the first generator because it reuses some templates from other generators.
  new TdpGenerator(),
  // Generators below does not have overlapping templates.
  new DefaultTemplateGenerator(),
  new OfficeAddinGeneratorNew(),
  new SPFxGeneratorNew(),
  new SPFxGeneratorImport(),
  new SsrTabGenerator(),
  new DeclarativeAgentWithExistingApiSpecGenerator(),
  new CustomEngineAgentWithExistingApiSpecGenerator(),
  new MessageExtensionWithExistingApiSpecGenerator(),
  new DeclarativeAgentGenerator(),
];
