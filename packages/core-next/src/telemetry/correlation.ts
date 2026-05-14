// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { AsyncLocalStorage } from "async_hooks";

const correlationStorage = new AsyncLocalStorage<string>();

/**
 * Run `fn` inside an AsyncLocalStorage scope tied to the given correlation ID.
 * Nested calls to `getCurrentCorrelationId()` within `fn` will return `id`.
 */
export function correlationScope<T>(id: string, fn: () => T): T {
  return correlationStorage.run(id, fn);
}

/**
 * Retrieve the current correlation ID from AsyncLocalStorage.
 * Returns empty string if called outside a `correlationScope`.
 */
export function getCurrentCorrelationId(): string {
  return correlationStorage.getStore() ?? "";
}
