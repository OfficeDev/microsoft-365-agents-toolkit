// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { SensitivityLabel } from "@microsoft/teamsfx-api";

export interface ListSensitivityCacheValue {
  labels: SensitivityLabel[];
  unixTimestamp: number;
}
