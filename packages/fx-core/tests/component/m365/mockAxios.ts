// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { AxiosInstance, AxiosRequestConfig, AxiosResponse } from "axios";

export class MockAxios {
  public axiosDeleteResponses: Record<string, unknown> = {};
  public axiosGetResponses: Record<string, unknown> = {};
  public axiosPostResponses: Record<string, unknown> = {};
  public axiosPutResponses: Record<string, unknown> = {};
  public data: Record<string, unknown> = {};
  public instance: AxiosInstance;

  constructor(private sandbox: sinon.SinonSandbox) {
    this.instance = {
      defaults: {
        headers: {
          common: {},
        },
      },
      interceptors: {
        request: {
          use: this.sandbox.stub(),
        },
        response: {
          use: this.sandbox.stub(),
        },
      },
      delete: <T = unknown, R = AxiosResponse<T>>(
        url: string,
        config?: AxiosRequestConfig
      ): Promise<R> => {
        const response = this.axiosDeleteResponses[url] as any;
        return response.message !== undefined
          ? Promise.reject(response)
          : Promise.resolve(response);
      },
      get: <T = unknown, R = AxiosResponse<T>>(url: string): Promise<R> => {
        const response = this.axiosGetResponses[url] as any;
        return response.message !== undefined
          ? Promise.reject(response)
          : Promise.resolve(response);
      },
      post: <T = unknown, R = AxiosResponse<T>>(
        url: string,
        data?: any,
        config?: AxiosRequestConfig
      ): Promise<R> => {
        const response = this.axiosPostResponses[url] as any;
        this.data[url] = data;
        return response.message !== undefined
          ? Promise.reject(response)
          : Promise.resolve(response);
      },
      put: <T = unknown, R = AxiosResponse<T>>(
        url: string,
        data?: any,
        config?: AxiosRequestConfig
      ): Promise<R> => {
        const response = this.axiosPutResponses[url] as any;
        return response.message !== undefined
          ? Promise.reject(response)
          : Promise.resolve(response);
      },
    } as unknown as AxiosInstance;
  }

  public reset(): void {
    this.axiosDeleteResponses = {};
    this.axiosGetResponses = {};
    this.axiosPostResponses = {};
    this.axiosPutResponses = {};
  }
}
