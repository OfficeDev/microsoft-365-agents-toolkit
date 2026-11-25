// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
"use strict";

import {
  ConfigFolderName,
  FxError,
  LogProvider,
  UserError,
  UserInteraction,
} from "@microsoft/teamsfx-api";
import "reflect-metadata";
import {
  BasicConstraintsExtension,
  ExtendedKeyUsage,
  ExtendedKeyUsageExtension,
  SubjectAlternativeNameExtension,
  X509Certificate,
  X509CertificateGenerator,
} from "@peculiar/x509";
import { createHash, createPrivateKey, createPublicKey, webcrypto } from "crypto";
import * as fs from "fs-extra";
import * as os from "os";
import * as path from "path";
import { v4 as uuidv4 } from "uuid";

import * as shellQuote from "shell-quote";
import { getDefaultString, getLocalizedString } from "../../common/localizeUtils";
import { CoreSource } from "../../error";
import { LocalDebugCertificate } from "./constants";
import * as ps from "./process";

export const localCertificateManagerDeps = {
  ensureDir: fs.ensureDir,
  pathExists: fs.pathExists,
  readFile: fs.readFile,
  writeFile: fs.writeFile,
  homedir: os.homedir,
  osType: os.type,
  execPowerShell: ps.execPowerShell,
  execShell: ps.execShell,
  localize: getLocalizedString,
};

const installText = () => localCertificateManagerDeps.localize("debug.install");
const learnMoreText = () => localCertificateManagerDeps.localize("core.provision.learnMore");
// TODO(xiaofhua): update help link for v3
const learnMoreUrl = "https://aka.ms/teamsfx-ca-certificate";
const warningMessage = () => localCertificateManagerDeps.localize("debug.warningMessage");
const confirmMessage = () =>
  warningMessage() + localCertificateManagerDeps.localize("debug.warningMessage2");

const trustCertificateCancelError = () =>
  new UserError({
    source: CoreSource,
    name: "TrustCertificateCancelError",
    helpLink: learnMoreUrl,
    message: getDefaultString("error.TrustCertificateCancelError"),
    displayMessage: localCertificateManagerDeps.localize("error.TrustCertificateCancelError"),
  });
export interface LocalCertificate {
  certPath: string;
  keyPath: string;
  isTrusted?: boolean;
  alreadyTrusted?: boolean;
  error?: FxError;
  found?: boolean;
}

export class LocalCertificateManager {
  private readonly ui?: UserInteraction;
  private readonly logger?: LogProvider;
  private readonly certFolder: string;

  constructor(ui?: UserInteraction, logger?: LogProvider) {
    this.ui = ui;
    this.logger = logger;
    this.certFolder = path
      .normalize(`${localCertificateManagerDeps.homedir()}/.${ConfigFolderName}/certificate`)
      .split(path.sep)
      .join(path.posix.sep);
  }

  /**
   * Local certificates are located at {home}/.fx/certificate/
   * Public certificate should be trusted into user"s certificate store.
   *
   * - Check and generate cert and key files (subject, usage, expiration, ...)
   * - Check cert store if trusted (thumbprint, expiration)
   * - Add to cert store if not trusted (friendly name as well)
   */
  public async setupCertificate(needTrust: boolean, checkOnly = false): Promise<LocalCertificate> {
    const certFilePath = `${this.certFolder}/${LocalDebugCertificate.CertFileName}`;
    const keyFilePath = `${this.certFolder}/${LocalDebugCertificate.KeyFileName}`;
    const localCert: LocalCertificate = {
      certPath: certFilePath,
      keyPath: keyFilePath,
    };

    try {
      let certThumbprint: string | undefined = undefined;
      await localCertificateManagerDeps.ensureDir(this.certFolder);

      if (
        (await localCertificateManagerDeps.pathExists(certFilePath)) &&
        (await localCertificateManagerDeps.pathExists(keyFilePath))
      ) {
        const certContent = await localCertificateManagerDeps.readFile(certFilePath, {
          encoding: "utf8",
        });
        const keyContent = await localCertificateManagerDeps.readFile(keyFilePath, {
          encoding: "utf8",
        });
        const verifyRes = this.verifyCertificateContent(certContent, keyContent);
        if (verifyRes[1]) {
          certThumbprint = verifyRes[0];
        }
      }

      localCert.found = !!certThumbprint;
      if (!certThumbprint) {
        if (checkOnly) {
          return localCert;
        }
        // generate cert and key
        certThumbprint = await this.generateCertificate(certFilePath, keyFilePath);
      }

      if (needTrust) {
        if (await this.verifyCertificateInStore(certThumbprint)) {
          // already trusted
          localCert.isTrusted = true;
          localCert.alreadyTrusted = true;
        } else {
          localCert.alreadyTrusted = false;
          if (checkOnly) {
            return localCert;
          }
          await this.trustCertificate(
            localCert,
            certThumbprint,
            LocalDebugCertificate.FriendlyName
          );
        }
      }
    } catch (error: any) {
      this.logger?.warning(`Failed to setup certificate. Error: ${error}`);
      localCert.isTrusted = false;
      localCert.error = new UserError({
        error,
        source: CoreSource,
        name: "SetupCertificateError",
        helpLink: learnMoreUrl,
      });
    } finally {
      return localCert;
    }
  }

  async generateCertificate(certFile: string, keyFile: string): Promise<string> {
    // prepare attributes and extensions
    const now = new Date();
    const expiry = new Date();
    expiry.setFullYear(expiry.getFullYear() + 1);

    const serialNumber = uuidv4().replace(/-/g, "");
    const keys = await webcrypto.subtle.generateKey(
      {
        name: "RSASSA-PKCS1-v1_5",
        modulusLength: 4096,
        publicExponent: new Uint8Array([1, 0, 1]),
        hash: "SHA-256",
      },
      true,
      ["sign", "verify"]
    );
    const cert = await X509CertificateGenerator.createSelfSigned(
      {
        serialNumber,
        name: "CN=localhost",
        notBefore: now,
        notAfter: expiry,
        keys,
        extensions: [
          new BasicConstraintsExtension(false),
          new ExtendedKeyUsageExtension([ExtendedKeyUsage.serverAuth]),
          new SubjectAlternativeNameExtension([{ type: "dns", value: "localhost" }]),
        ],
      },
      globalThis.crypto
    );

    // SHA-1 identifies certificates in OS trust stores; it is not used for signing.
    const thumbprint = createHash("sha1").update(Buffer.from(cert.rawData)).digest("hex");

    // output
    const certContent = cert.toString("pem");
    const privateKey = createPrivateKey({
      key: Buffer.from(await webcrypto.subtle.exportKey("pkcs8", keys.privateKey)),
      format: "der",
      type: "pkcs8",
    });
    const keyContent = privateKey.export({ format: "pem", type: "pkcs1" }).toString();
    await localCertificateManagerDeps.writeFile(certFile, certContent, { encoding: "utf8" });
    await localCertificateManagerDeps.writeFile(keyFile, keyContent, { encoding: "utf8" });

    return thumbprint;
  }

  verifyCertificateContent(certContent: string, keyContent: string): [string | undefined, boolean] {
    const thumbprint: string | undefined = undefined;
    try {
      const cert = new X509Certificate(certContent);
      const privateKey = createPrivateKey(keyContent);

      // get thumbprint
      const thumbprint = createHash("sha1").update(Buffer.from(cert.rawData)).digest("hex");

      // verify key pair
      if (
        privateKey.asymmetricKeyType !== "rsa" ||
        !createPublicKey(privateKey)
          .export({ format: "der", type: "spki" })
          .equals(Buffer.from(cert.publicKey.rawData))
      ) {
        return [thumbprint, false];
      }

      // verify subject and issuer
      if (cert.subjectName.getField("CN")[0] !== "localhost") {
        return [thumbprint, false];
      }

      if (cert.issuerName.getField("CN")[0] !== "localhost") {
        return [thumbprint, false];
      }

      // verify date, add one day buffer
      const now = new Date();
      const tomorrow = new Date();
      tomorrow.setDate(tomorrow.getDate() + 1);
      const notBefore = cert.notBefore;
      const notAfter = cert.notAfter;
      if (notBefore > now || notAfter < tomorrow) {
        return [thumbprint, false];
      }

      // verify extension
      const basicConstraints = cert.getExtension(BasicConstraintsExtension);
      if (!basicConstraints || basicConstraints.ca) {
        return [thumbprint, false];
      }

      const extKeyUsage = cert.getExtension(ExtendedKeyUsageExtension);
      if (!extKeyUsage?.usages.includes(ExtendedKeyUsage.serverAuth)) {
        return [thumbprint, false];
      }

      const subjectAltName = cert.getExtension(SubjectAlternativeNameExtension);
      if (
        !subjectAltName?.names.items.some(
          (name) => name.type === "dns" && name.value === "localhost"
        )
      ) {
        return [thumbprint, false];
      }

      return [thumbprint, true];
    } catch (error) {
      // treat any error as not verified, to not block the main progress
      return [thumbprint, false];
    }
  }

  async verifyCertificateInStore(thumbprint: string): Promise<boolean | undefined> {
    try {
      if (localCertificateManagerDeps.osType() === "Windows_NT") {
        return await this.checkCertificateWindows(thumbprint);
      } else if (localCertificateManagerDeps.osType() === "Darwin") {
        const listCertCommand = `security find-certificate -c localhost -a -Z -p "${localCertificateManagerDeps.homedir()}/Library/Keychains/login.keychain-db"`;
        const existingCertificates = await localCertificateManagerDeps.execShell(listCertCommand);
        if (existingCertificates) {
          const thumbprintRegex = /SHA-1 hash: ([0-9A-Z]+)/g;
          let match = undefined;
          while ((match = thumbprintRegex.exec(existingCertificates))) {
            const existingThumbprint = match[1];
            if (existingThumbprint.toUpperCase() === thumbprint.toUpperCase()) {
              return true;
            }
          }
        }

        return false;
      } else {
        // TODO: Linux
        return undefined;
      }
    } catch (error) {
      // treat any error as not verified, to not block the main progress

      this.logger?.debug(`Certificate unverified. Details: ${error}`);
      return false;
    }
  }

  private async trustCertificate(
    localCert: LocalCertificate,
    thumbprint: string,
    friendlyName: string
  ): Promise<void> {
    try {
      if (localCertificateManagerDeps.osType() === "Windows_NT") {
        if (!(await this.waitForUserConfirm())) {
          localCert.isTrusted = false;
          localCert.error = trustCertificateCancelError();
          return;
        }

        await this.trustCertificateWindows(localCert, thumbprint, friendlyName);

        localCert.isTrusted = true;
        return;
      } else if (localCertificateManagerDeps.osType() === "Darwin") {
        if (!(await this.waitForUserConfirm())) {
          localCert.isTrusted = false;
          localCert.error = trustCertificateCancelError();
          return;
        }

        await localCertificateManagerDeps.execShell(
          `security add-trusted-cert -p ssl -k "${localCertificateManagerDeps.homedir()}/Library/Keychains/login.keychain-db" "${
            localCert.certPath
          }"`
        );

        localCert.isTrusted = true;
        return;
      } else {
        // TODO: Linux
        localCert.isTrusted = undefined;
        return;
      }
    } catch (error: any) {
      // treat any error as install failure, to not block the main progress

      this.logger?.warning(`Failed to install certificate. Error: ${error}`);
      localCert.isTrusted = false;
      localCert.error = new UserError({
        error,
        source: CoreSource,
        name: "TrustCertificateError",
        helpLink: learnMoreUrl,
      });
      return;
    }
  }

  private async waitForUserConfirm(): Promise<boolean> {
    if (this.ui) {
      let userSelected: string | undefined;
      do {
        const res = await this.ui.showMessage(
          "info",
          confirmMessage(),
          true,
          learnMoreText(),
          installText()
        );
        userSelected = res.isOk() ? res.value : undefined;
        if (userSelected === learnMoreText()) {
          void this.ui.openUrl(learnMoreUrl);
        }
      } while (userSelected === learnMoreText());
      return userSelected === installText();
    }

    // No dialog, always return true;
    return true;
  }

  private async checkCertificateWindows(thumbprint: string): Promise<boolean> {
    const quotedThumbprint: string = shellQuote.quote([thumbprint]);
    try {
      // try powershell first
      const getCertCommand = `Get-ChildItem -Path Cert:\\CurrentUser\\Root | Where-Object { $_.Thumbprint -match '${quotedThumbprint}' }`;
      const getCertRes = await localCertificateManagerDeps.execPowerShell(getCertCommand);
      return getCertRes.toUpperCase().includes(thumbprint.toUpperCase());
    } catch (error: any) {
      // if any error, try certutil
      const getCertCommand = `certutil -user -verifystore root ${quotedThumbprint}`;
      const getCertRes = (await localCertificateManagerDeps.execShell(getCertCommand)).trim();
      return getCertRes.toUpperCase().includes(thumbprint.toUpperCase());
    }
  }

  private async trustCertificateWindows(
    localCert: LocalCertificate,
    thumbprint: string,
    friendlyName: string
  ): Promise<void> {
    try {
      // try powershell first
      const installCertCommand = `Import-Certificate -FilePath '${localCert.certPath}' -CertStoreLocation Cert:\\CurrentUser\\Root`;
      await localCertificateManagerDeps.execPowerShell(installCertCommand);
      try {
        const friendlyNameCommand = `(Get-ChildItem -Path Cert:\\CurrentUser\\Root\\${thumbprint}).FriendlyName='${friendlyName}'`;
        await localCertificateManagerDeps.execPowerShell(friendlyNameCommand);
      } catch (e) {
        // ignore friendly name failure
      }
    } catch (error: any) {
      // if any error, try certutil
      const installCertCommand = `certutil -user -addstore root "${localCert.certPath}"`;
      await localCertificateManagerDeps.execShell(installCertCommand);
      try {
        const certInfPath = path.join(path.dirname(localCert.certPath), "localhost.inf");
        await localCertificateManagerDeps.writeFile(
          certInfPath,
          [
            "[Version]",
            `Signature = "$Windows NT$"`,
            "[Properties]",
            `11 = {text}${friendlyName}`,
          ].join(os.EOL)
        );
        const friendlyNameCommand = `certutil -user -repairstore root ${thumbprint} "${certInfPath}"`;
        await localCertificateManagerDeps.execShell(friendlyNameCommand);
      } catch (e) {
        // ignore friendly name failure
      }
    }
  }
}
