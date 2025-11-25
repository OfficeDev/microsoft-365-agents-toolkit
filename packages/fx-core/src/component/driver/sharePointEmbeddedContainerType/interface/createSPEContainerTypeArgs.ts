// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { SPContainerTypeBillingClassification } from "./sharePointEmbeddedContainerType";

export interface CreateSPEContainerTypeArgs {
  owningApplicationId: string; // The application ID of the Microsoft Entra app that will own the container type;
  billingClassification: SPContainerTypeBillingClassification; // Billing classification for the container type;
  name?: string; // The name for the container type;
  discoverable?: boolean; // Whether the container type is discoverable by M365 apps, including Copilot;
}
