// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { setKiotaConfig, searchDescription, getKiotaTree, generatePlugin } from "@microsoft/kiota";

// Re-export the kiota functions as a module that can be stubbed in tests
export const kiota = {
  setKiotaConfig,
  searchDescription,
  getKiotaTree,
  generatePlugin,
};
