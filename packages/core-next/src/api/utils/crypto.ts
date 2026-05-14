// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { Result } from "neverthrow";
import { FxError } from "../error";

export interface CryptoProvider {
  encrypt(plaintext: string): Result<string, FxError>;
  decrypt(ciphertext: string): Result<string, FxError>;
}
