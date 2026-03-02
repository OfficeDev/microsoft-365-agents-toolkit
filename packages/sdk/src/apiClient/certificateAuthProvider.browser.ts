// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { AxiosRequestConfig } from "axios";
import { SecureContextOptions } from "tls";
import { AuthProvider } from "./authProvider";
import { ErrorWithCode, ErrorCode, ErrorMessage } from "../core/errors";
import { formatString } from "../util/utils";

/**
 * Provider that handles Certificate authentication
 * @deprecated This package will be deprecated by 2026-07. Please use [Microsoft 365 Agents SDK](https://www.npmjs.com/package/@microsoft/agents-hosting) instead.
 */

export class CertificateAuthProvider implements AuthProvider {
  private certOption: SecureContextOptions;

  /**
   *
   * @param { SecureContextOptions } certOption - information about the cert used in http requests
   * @deprecated This package will be deprecated by 2026-07. Please use [Microsoft 365 Agents SDK](https://www.npmjs.com/package/@microsoft/agents-hosting) instead.
   */
  constructor(certOption: SecureContextOptions) {
    throw new ErrorWithCode(
      formatString(ErrorMessage.BrowserRuntimeNotSupported, "CertificateAuthProvider"),
      ErrorCode.RuntimeNotSupported
    );
  }

  /**
   * Adds authentication info to http requests.
   *
   * @param { AxiosRequestConfig } config - Contains all the request information and can be updated to include extra authentication info.
   * Refer https://axios-http.com/docs/req_config for detailed document.
   *
   * @returns Updated axios request config.
   *
   * @throws {@link ErrorCode|InvalidParameter} - when custom httpsAgent in the request has duplicate properties with certOption provided in constructor.
   * @throws {@link ErrorCode|RuntimeNotSupported} when runtime is browser.
   * @deprecated This package will be deprecated by 2026-07. Please use [Microsoft 365 Agents SDK](https://www.npmjs.com/package/@microsoft/agents-hosting) instead.
   */
  public AddAuthenticationInfo(config: AxiosRequestConfig): Promise<AxiosRequestConfig> {
    return Promise.reject(
      new ErrorWithCode(
        formatString(ErrorMessage.BrowserRuntimeNotSupported, "CertificateAuthProvider"),
        ErrorCode.RuntimeNotSupported
      )
    );
  }
}

/**
 * Helper to create SecureContextOptions from PEM format cert
 *
 * @param { string | Buffer } cert - The cert chain in PEM format
 * @param { string | Buffer } key - The private key for the cert chain
 * @param { {passphrase?: string; ca?: string | Buffer} } options - Optional settings when create the cert options.
 *
 * @returns Instance of SecureContextOptions
 *
 * @throws {@link ErrorCode|InvalidParameter} - when any parameter is empty
 * @throws {@link ErrorCode|RuntimeNotSupported} when runtime is browser.
 * @deprecated This package will be deprecated by 2026-07. Please use [Microsoft 365 Agents SDK](https://www.npmjs.com/package/@microsoft/agents-hosting) instead.
 *
 */
export function createPemCertOption(
  cert: string | Buffer,
  key: string | Buffer,
  options?: {
    passphrase?: string;
    ca?: string | Buffer;
  }
): SecureContextOptions {
  throw new ErrorWithCode(
    formatString(ErrorMessage.BrowserRuntimeNotSupported, "createPemCertOption"),
    ErrorCode.RuntimeNotSupported
  );
}

/**
 * Helper to create SecureContextOptions from PFX format cert
 *
 * @param { string | Buffer } pfx - The content of .pfx file
 * @param { {passphrase?: string} } options - Optional settings when create the cert options.
 *
 * @returns Instance of SecureContextOptions
 *
 * @throws {@link ErrorCode|InvalidParameter} - when any parameter is empty
 * @throws {@link ErrorCode|RuntimeNotSupported} when runtime is browser.
 * @deprecated This package will be deprecated by 2026-07. Please use [Microsoft 365 Agents SDK](https://www.npmjs.com/package/@microsoft/agents-hosting) instead.
 *
 */
export function createPfxCertOption(
  pfx: string | Buffer,
  options?: {
    passphrase?: string;
  }
): SecureContextOptions {
  throw new ErrorWithCode(
    formatString(ErrorMessage.BrowserRuntimeNotSupported, "createPfxCertOption"),
    ErrorCode.RuntimeNotSupported
  );
}
