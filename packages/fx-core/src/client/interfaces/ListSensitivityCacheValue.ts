// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

export interface ListSensitivityCacheValue {
  labels: SensitivityLabel[];
  unixTimestamp: number;
}

export class SensitivityLabel {
  id?: string;
  name?: string;
  description?: string;
  displayName?: string;
}
