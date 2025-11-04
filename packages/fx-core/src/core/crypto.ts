// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { CryptoProvider, err, FxError, ok, Result, SystemError } from "@microsoft/teamsfx-api";
import Cryptr from "cryptr";

export class LocalCrypto implements CryptoProvider {
  private cryptr: Cryptr;
  private fixedCryptr: Cryptr;
  private prefix = "crypto_";

  constructor(projectId: string) {
    this.cryptr = new Cryptr(projectId + "_teamsfx");
    this.fixedCryptr = new Cryptr("teamsfx_global_key");
  }

  public encrypt(plaintext: string): Result<string, FxError> {
    return ok(this.prefix + this.fixedCryptr.encrypt(plaintext));
  }

  public decrypt(ciphertext: string): Result<string, FxError> {
    if (!ciphertext.startsWith(this.prefix)) {
      // legacy raw secret string
      return ok(ciphertext);
    }
    const encryptedData = ciphertext.substr(this.prefix.length);
    try {
      return ok(this.fixedCryptr.decrypt(encryptedData));
    } catch (e) {
      try {
        return ok(this.cryptr.decrypt(encryptedData));
      } catch (e2) {
        return err(new SystemError("Core", "DecryptionError", "Cipher text is broken"));
      }
    }
  }
}
