// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { SignInAudience } from "./signInAudience";

export interface CreateAadAppArgs {
  name: string; // The name of AAD app
  generateClientSecret: boolean; // Whether generate client secret for the app
  generateServicePrincipal?: boolean; // Whether generate service principal for the app
  signInAudience?: SignInAudience; // Specifies what Microsoft accounts are supported for the current application.
  clientSecretExpireDays?: number; // The number of days the client secret is valid
  clientSecretDescription?: string; // The description of the client secret
  serviceManagementReference?: string; // Used as service tree id
}
