// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import axios, { AxiosInstance, AxiosResponse, InternalAxiosRequestConfig } from "axios";
import { AtkContext } from "../core/context";
import { maskSecret } from "../secretMasker/masker";

export interface HttpClientOptions {
  /** Base URL for all requests */
  baseURL?: string;
  /** Default timeout in milliseconds */
  timeout?: number;
  /** Additional default headers */
  headers?: Record<string, string>;
}

/**
 * Create an Axios instance wired with telemetry interceptors.
 * Each request/response is logged via ctx.logger and ctx.telemetry.
 * Sensitive values in URLs and headers are masked before telemetry.
 *
 * @param ctx - AtkContext for logging and telemetry
 * @param options - Optional Axios configuration overrides
 */
export function createHttpClient(ctx: AtkContext, options?: HttpClientOptions): AxiosInstance {
  const instance = axios.create({
    baseURL: options?.baseURL,
    timeout: options?.timeout,
    headers: options?.headers,
  });

  // Request interceptor — log start
  instance.interceptors.request.use((config: InternalAxiosRequestConfig) => {
    const safeUrl = maskSecret(config.url ?? "");
    ctx.logger.debug(`[http] ${config.method?.toUpperCase()} ${safeUrl}`);
    (config as InternalAxiosRequestConfig & { _startTime: number })._startTime = Date.now();
    return config;
  });

  // Response interceptor — log success and error
  instance.interceptors.response.use(
    (response: AxiosResponse) => {
      const startTime = (response.config as InternalAxiosRequestConfig & { _startTime?: number })
        ._startTime;
      const duration = startTime ? Date.now() - startTime : 0;
      const safeUrl = maskSecret(response.config.url ?? "");
      ctx.logger.debug(`[http] ${response.status} ${safeUrl} (${duration}ms)`);
      ctx.telemetry.sendTelemetryEvent(
        "http-request",
        {
          method: response.config.method?.toUpperCase() ?? "GET",
          status: String(response.status),
          url: safeUrl,
        },
        { duration }
      );
      return response;
    },
    (error) => {
      if (axios.isAxiosError(error) && error.config) {
        const startTime = (error.config as InternalAxiosRequestConfig & { _startTime?: number })
          ._startTime;
        const duration = startTime ? Date.now() - startTime : 0;
        const safeUrl = maskSecret(error.config.url ?? "");
        const status = error.response?.status ?? 0;
        ctx.logger.error(`[http] ${status} ${safeUrl} (${duration}ms) - ${error.message}`);
        ctx.telemetry.sendTelemetryErrorEvent(
          "http-request",
          {
            method: error.config.method?.toUpperCase() ?? "GET",
            status: String(status),
            url: safeUrl,
          },
          { duration }
        );
      }
      return Promise.reject(error);
    }
  );

  return instance;
}
