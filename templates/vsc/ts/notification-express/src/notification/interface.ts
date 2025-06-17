// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { Storage, StoreItems } from "@microsoft/agents-hosting";

export interface IStorage extends Storage {
  list(pageSize?: number, continuationToken?: string): Promise<PagedData<StoreItems>>;
}

/**
 * Represents a page of data.
 */
export interface PagedData<T> {
  /**
   * Page of data.
   */
  data: T[];

  /**
   * The Continuation Token to pass to get the next page of results.
   *
   * @remarks
   * Undefined or empty token means the page reaches the end.
   */
  continuationToken?: string;
}
