// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

export type CreateSPEContainerTypeOutput = {
  containerTypeId?: string;
};

// The const is used to reference the property name in CreateAadAppOutput. When renaming the properties in CreateAadAppOutput, you need to update the const as well.
export const OutputKeys = {
  containerTypeId: "containerTypeId",
};
