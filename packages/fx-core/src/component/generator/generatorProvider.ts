// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import { SpecGenerator } from "./apiSpec/generator";
import { CopilotExtensionGenerator } from "./copilotExtension/generator";
import { DefaultTemplateGenerator } from "./defaultGenerator";
import { OfficeAddinGeneratorNew } from "./officeAddin/generator";
import { SsrTabGenerator } from "./other/ssrTabGenerator";
import { TdpGenerator } from "./other/tdpGenerator";
import { SPFxGeneratorImport, SPFxGeneratorNew } from "./spfx/spfxGenerator";

// When multiple generators are activated, only the top one will be executed.
export const Generators = [
  new TdpGenerator(),
  new OfficeAddinGeneratorNew(),
  new SsrTabGenerator(),
  new DefaultTemplateGenerator(),
  new SPFxGeneratorNew(),
  new SPFxGeneratorImport(),
  new SpecGenerator(),
  new CopilotExtensionGenerator(),
];
