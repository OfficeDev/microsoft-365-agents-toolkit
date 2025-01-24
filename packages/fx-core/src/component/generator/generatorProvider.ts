// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import { SpecGenerator } from "./apiSpec/generator";
import { CopilotExtensionGenerator } from "./copilotExtension/generator";
import { OfficeAddinGeneratorNew } from "./officeAddin/generator";
import { SPFxGeneratorImport, SPFxGeneratorNew } from "./spfx/spfxGenerator";
import { TdpGenerator } from "./TdpGenerator";
import { SsrTabGenerator } from "./templates/ssrTabGenerator";
import { DefaultTemplateGenerator } from "./templates/templateGenerator";

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
