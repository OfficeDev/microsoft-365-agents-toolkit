// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

export interface ExpServiceProvider {
  getTreatmentVariableAsync<T extends boolean | number | string>(
    configId: string,
    name: string,
    checkCache?: boolean
  ): Promise<T | undefined>;
}
