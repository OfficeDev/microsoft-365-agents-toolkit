// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author quke@microsoft.com
 */

// hello-world-office-addin deploys to Azure Static Web Apps. It is an Office
// Add-in (not a Teams app), so validate, package, and preview steps (which
// expect a Teams app manifest) are skipped.

import { TemplateProjectFolder } from "../../utils/constants";
import { CaseFactory } from "./sampleCaseFactory";

class HelloWorldOfficeAddinTestCase extends CaseFactory {}

new HelloWorldOfficeAddinTestCase(
  TemplateProjectFolder.HelloWorldOfficeAddin,
  0,
  "quke@microsoft.com",
  [],
  { skipValidate: true, skipPackage: true, skipPreview: true }
).test();
