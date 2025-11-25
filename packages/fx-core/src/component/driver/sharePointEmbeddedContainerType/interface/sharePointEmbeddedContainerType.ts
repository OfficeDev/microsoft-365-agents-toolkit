// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

export enum SPContainerTypeBillingClassification {
  standard = "standard", // standard container type.
  directToCustomer = "directToCustomer", // direct to customer container type.
  trial = "trial", // trial container type.
}

export enum SPContainerBillingStatus {
  invalid,
  valid,
}

export enum SPContainerTypeSettingsOverride {
  none,
  isItemVersioningEnabled,
  itemMajorVersionLimit,
  maxStoragePerContainerInBytes,
  unknownFutureValue,
}

export interface ISharePointEmbeddedContainerType {
  id?: string; // container type ID;
  name?: string; // container type name;
  owningAppId?: string; // The application ID of the Microsoft Entra app that owns the container type;
  billingClassification?: SPContainerTypeBillingClassification; // Billing classification for the container type;
  billingStatus?: SPContainerBillingStatus; // Billing status for the container type;
  createdDateTime?: Date; // The date and time when the container type was created;
  expirationDateTime?: Date; // The date and time when the container type will expire (only if it's trial container type);
  settings?: SharePointEmbeddedContainerTypeSettings; // Settings for the container type;
  etag?: string; //Used in update for optimistic concurrency control.
}

export interface SharePointEmbeddedContainerTypeSettings {
  sharingCapability?: string; // sharing capabilities permitted for containers.
  urlTemplate?: string; // Pattern used to redirect files
  isDiscoverabilityEnabled?: boolean; // Enables, disables surface of items from containers in experiences like my activity or M356. Optional.
  isSearchEnabled?: boolean; // If search is enabled. Optional.
  isItemVersioningEnabled?: boolean; // Controls item versioning. Optional.
  itemMajorVersionLimit?: number; // Maximum number of versions. Versioning must be enabled. Optional.
  maxStoragePerContainerInBytes?: number; // Controls maximum storage in bytes. Optional.
  isSharingRestricted?: boolean; // Controls if sharing is restricted. Optional.
  consumingTenantOverridables?: SPContainerTypeSettingsOverride; // Settings that can be overwritten in the consuming tenant, comma separated. Optional.
}
