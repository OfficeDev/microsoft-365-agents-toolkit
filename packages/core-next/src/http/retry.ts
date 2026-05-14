// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { AxiosResponse } from "axios";

const DEFAULT_RETRIES = 3;
const RETRY_DELAY_MS = 1000;

/**
 * Retry an HTTP request function on server errors (5xx) or missing status.
 *
 * @param fn - Function that performs the HTTP request
 * @param retries - Number of retry attempts (default: 3)
 */
export async function sendWithRetry<T>(
  fn: () => Promise<AxiosResponse<T>>,
  retries: number = DEFAULT_RETRIES
): Promise<AxiosResponse<T>> {
  let lastError: unknown;
  for (let attempt = 0; attempt <= retries; attempt++) {
    try {
      const response = await fn();
      if (response.status && (response.status < 500 || response.status >= 600)) {
        return response;
      }
      // 5xx — will retry
      lastError = new Error(`Server error: ${response.status}`);
    } catch (error) {
      // Don't retry client errors (4xx) except retryable ones:
      //  - 429 Too Many Requests (rate-limiting)
      //  - 412 Precondition Failed (transient etag / propagation races)
      if (error && typeof error === "object" && "response" in error) {
        const status = (error as any).response?.status;
        if (
          typeof status === "number" &&
          status >= 400 &&
          status < 500 &&
          status !== 429 &&
          status !== 412
        ) {
          throw error;
        }
      }
      lastError = error;
    }
    if (attempt < retries) {
      await delay(RETRY_DELAY_MS * (attempt + 1));
    }
  }
  throw lastError;
}

/**
 * Execute an HTTP request with a timeout and optional retries.
 * Uses AbortController for clean cancellation (replaces old CancelToken pattern).
 *
 * @param fn - Function that receives an AbortSignal and performs the request
 * @param timeoutMs - Timeout in milliseconds
 * @param retries - Number of retry attempts (default: 0)
 */
export async function sendWithTimeout<T>(
  fn: (signal: AbortSignal) => Promise<AxiosResponse<T>>,
  timeoutMs: number,
  retries: number = 0
): Promise<AxiosResponse<T>> {
  const attempt = async (): Promise<AxiosResponse<T>> => {
    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), timeoutMs);
    try {
      return await fn(controller.signal);
    } finally {
      clearTimeout(timer);
    }
  };

  if (retries > 0) {
    return sendWithRetry(attempt, retries);
  }
  return attempt();
}

function delay(ms: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, ms));
}
