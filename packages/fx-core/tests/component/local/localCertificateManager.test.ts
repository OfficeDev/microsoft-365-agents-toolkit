// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { ConfigFolderName, FxError, Result, UserInteraction, ok } from "@microsoft/teamsfx-api";
import "reflect-metadata";
import {
  BasicConstraintsExtension,
  ExtendedKeyUsage,
  ExtendedKeyUsageExtension,
  Extension,
  SubjectAlternativeNameExtension,
  X509Certificate,
  X509CertificateCreateCommonParams,
  X509CertificateGenerator,
} from "@peculiar/x509";
import {
  createHash,
  createPrivateKey,
  webcrypto,
  X509Certificate as NodeCertificate,
} from "crypto";
import { once } from "events";
import fs from "fs-extra";
import { createServer, get } from "https";
import * as path from "path";
import { chai, vi } from "vitest";
import {
  LocalCertificateManager,
  localCertificateManagerDeps,
} from "../../../src/component/local/localCertificateManager";

let lightweightCertPair: { certPem: string; keyPem: string; thumbprint: string };

beforeAll(async () => {
  lightweightCertPair = await createLightweightCertPair();
});

describe("certificate", () => {
  const workspaceFolder = path.resolve(__dirname, "../data/n/t/r/test space 1/");
  const expectedWorkspaceFolder = path.resolve(__dirname, "../data/n/t/r/test space 1/");
  const expectedCertFile = path.resolve(
    expectedWorkspaceFolder,
    `.home/.${ConfigFolderName}/certificate/localhost.crt`
  );
  const expectedKeyFile = path.resolve(
    expectedWorkspaceFolder,
    `.home/.${ConfigFolderName}/certificate/localhost.key`
  );
  describe("setupCertificate", () => {
    const fakeHomeDir = path.resolve(workspaceFolder, ".home/");
    let files: Record<string, any> = {};
    let certManager: LocalCertificateManager;

    beforeEach(() => {
      files = {};
      vi.restoreAllMocks();
      vi.spyOn(LocalCertificateManager.prototype, "generateCertificate").mockImplementation(
        async (certFile: string, keyFile: string) => {
          files[path.resolve(certFile)] = lightweightCertPair.certPem;
          files[path.resolve(keyFile)] = lightweightCertPair.keyPem;
          return lightweightCertPair.thumbprint;
        }
      );
      vi.spyOn(localCertificateManagerDeps, "ensureDir").mockImplementation(async (dir: string) => {
        return Promise.resolve();
      });
      vi.spyOn(localCertificateManagerDeps, "pathExists").mockImplementation(
        async (file: string) => {
          return Promise.resolve(files[path.resolve(file)] !== undefined);
        }
      );
      vi.spyOn(localCertificateManagerDeps, "readFile").mockImplementation(
        async (file: fs.PathLike | number, options?: any) => {
          return Promise.resolve(files[path.resolve(file as string)]);
        }
      );
      vi.spyOn(localCertificateManagerDeps, "writeFile").mockImplementation(
        async (file: fs.PathLike | number, data: any, options?: any) => {
          files[path.resolve(file as string)] = data;
          return Promise.resolve();
        }
      );
      vi.spyOn(localCertificateManagerDeps, "homedir").mockImplementation(() => fakeHomeDir);
      vi.spyOn(localCertificateManagerDeps, "execPowerShell").mockImplementation(
        async (command: string) => {
          if (command.startsWith("Get-ChildItem")) {
            // Command: `Get-ChildItem -Path Cert:\\CurrentUser\\Root | Where-Object { $_.Thumbprint -match '${thumbprint}' }`
            return command.split("'")[1];
          } else if (command.startsWith("Import-Certificate")) {
            // Command: `Import-Certificate -FilePath '${localCert.certPath}' -CertStoreLocation Cert:\\CurrentUser\\Root)`
            return "thumbprint";
          } else {
            return "";
          }
        }
      );
      certManager = new LocalCertificateManager();
    });

    afterEach(() => {
      vi.restoreAllMocks();
    });

    [
      { osType: "Windows_NT", isTrusted: true },
      { osType: "Linux", isTrusted: undefined },
    ].forEach((data) => {
      it(`happy path ${data.osType}`, async () => {
        vi.spyOn(localCertificateManagerDeps, "osType").mockReturnValue(data.osType);
        const res = await certManager.setupCertificate(true);

        chai.assert.equal(
          res.certPath,
          path.normalize(expectedCertFile).split(path.sep).join(path.posix.sep)
        );
        chai.assert.equal(
          res.keyPath,
          path.normalize(expectedKeyFile).split(path.sep).join(path.posix.sep)
        );

        const certContent = files[path.resolve(expectedCertFile)];
        chai.assert.isDefined(certContent);
        chai.assert.isTrue(
          /-----BEGIN CERTIFICATE-----.*-----END CERTIFICATE-----/gs.test(certContent)
        );
        const keyContent = files[path.resolve(expectedKeyFile)];
        chai.assert.isDefined(keyContent);
        chai.assert.isTrue(
          /-----BEGIN RSA PRIVATE KEY-----.*-----END RSA PRIVATE KEY-----/gs.test(keyContent)
        );
        chai.assert.equal(res.isTrusted, data.isTrusted);
      });
    });

    [
      { osType: "Windows_NT", isTrusted: undefined },
      { osType: "Linux", isTrusted: undefined },
    ].forEach((data) => {
      it(`skip trust ${data.osType}`, async () => {
        vi.spyOn(localCertificateManagerDeps, "osType").mockReturnValue(data.osType);
        const res = await certManager.setupCertificate(false);

        const certContent = files[path.resolve(expectedCertFile)];
        chai.assert.isDefined(certContent);
        chai.assert.isTrue(
          /-----BEGIN CERTIFICATE-----.*-----END CERTIFICATE-----/gs.test(certContent)
        );
        const keyContent = files[path.resolve(expectedKeyFile)];
        chai.assert.isDefined(keyContent);
        chai.assert.isTrue(
          /-----BEGIN RSA PRIVATE KEY-----.*-----END RSA PRIVATE KEY-----/gs.test(keyContent)
        );
        chai.assert.equal(res.isTrusted, data.isTrusted);
      });
    });

    [
      { osType: "Windows_NT", isTrusted: true },
      { osType: "Linux", isTrusted: undefined },
    ].forEach((data) => {
      it(`existing verified cert ${data.osType}`, async () => {
        vi.spyOn(localCertificateManagerDeps, "osType").mockReturnValue(data.osType);
        let res = await certManager.setupCertificate(true);
        const certContent1 = files[path.resolve(expectedCertFile)];
        chai.assert.isDefined(certContent1);
        const thumbprint1 = getCertThumbprint(certContent1);

        res = await certManager.setupCertificate(true);
        const certContent2 = files[path.resolve(expectedCertFile)];
        chai.assert.isDefined(certContent2);
        const keyContent = files[path.resolve(expectedKeyFile)];
        chai.assert.isDefined(keyContent);
        const thumbprint2 = getCertThumbprint(certContent2);
        chai.assert.equal(thumbprint1, thumbprint2);
        chai.assert.equal(res.isTrusted, data.isTrusted);
      });
    });
  });

  describe("setupCertificate certutil", () => {
    const fakeHomeDir = path.resolve(workspaceFolder, ".home/");
    let files: Record<string, any> = {};
    let certManager: LocalCertificateManager;

    beforeEach(() => {
      files = {};
      vi.restoreAllMocks();
      vi.spyOn(LocalCertificateManager.prototype, "generateCertificate").mockImplementation(
        async (certFile: string, keyFile: string) => {
          files[path.resolve(certFile)] = lightweightCertPair.certPem;
          files[path.resolve(keyFile)] = lightweightCertPair.keyPem;
          return lightweightCertPair.thumbprint;
        }
      );
      vi.spyOn(localCertificateManagerDeps, "osType").mockReturnValue("Windows_NT");
      vi.spyOn(localCertificateManagerDeps, "ensureDir").mockResolvedValue();
      vi.spyOn(localCertificateManagerDeps, "pathExists").mockImplementation(
        async (file: string) => {
          return Promise.resolve(files[path.resolve(file)] !== undefined);
        }
      );
      vi.spyOn(localCertificateManagerDeps, "readFile").mockImplementation(
        async (file: fs.PathLike | number, options?: any) => {
          return Promise.resolve(files[path.resolve(file as string)]);
        }
      );
      vi.spyOn(localCertificateManagerDeps, "writeFile").mockImplementation(
        async (file: fs.PathLike | number, data: any, options?: any) => {
          files[path.resolve(file as string)] = data;
          return Promise.resolve();
        }
      );
      vi.spyOn(localCertificateManagerDeps, "homedir").mockImplementation(() => fakeHomeDir);
      vi.spyOn(localCertificateManagerDeps, "execPowerShell").mockRejectedValue();
      vi.spyOn(localCertificateManagerDeps, "execShell").mockImplementation(
        async (command: string) => {
          if (command.startsWith("certutil -user -verifystore")) {
            // Command: `certutil -user -verifystore root ${thumbprint}`
            return "Not Found";
          } else if (command.startsWith("certutil -user -addstore")) {
            // Command: `certutil -user -addstore root "${localCert.certPath}"`
            return "addstore";
          } else if (command.startsWith("certutil -user -repairstore")) {
            // Command: `certutil -user -repairstore root ${thumbprint} "${certInfPath}"`
            return "repairstore";
          } else {
            return "";
          }
        }
      );
      certManager = new LocalCertificateManager();
    });

    afterEach(() => {
      vi.restoreAllMocks();
    });

    it(`happy path windows`, async () => {
      const res = await certManager.setupCertificate(true);

      chai.assert.equal(
        res.certPath,
        path.normalize(expectedCertFile).split(path.sep).join(path.posix.sep)
      );
      chai.assert.equal(
        res.keyPath,
        path.normalize(expectedKeyFile).split(path.sep).join(path.posix.sep)
      );

      const certContent = files[path.resolve(expectedCertFile)];
      chai.assert.isDefined(certContent);
      chai.assert.isTrue(
        /-----BEGIN CERTIFICATE-----.*-----END CERTIFICATE-----/gs.test(certContent)
      );
      const keyContent = files[path.resolve(expectedKeyFile)];
      chai.assert.isDefined(keyContent);
      chai.assert.isTrue(
        /-----BEGIN RSA PRIVATE KEY-----.*-----END RSA PRIVATE KEY-----/gs.test(keyContent)
      );
      chai.assert.equal(res.isTrusted, true);
    });
  });

  describe("platform specific", () => {
    const fakeHomeDir = path.resolve(workspaceFolder, ".home/");
    const files: Record<string, any> = {};
    let certManager: LocalCertificateManager;

    beforeEach(() => {
      vi.restoreAllMocks();
    });

    afterEach(() => {
      vi.restoreAllMocks();
    });

    it("waitForUserConfirm once", async () => {
      vi.spyOn(localCertificateManagerDeps, "localize").mockImplementation((key, ...params) => {
        if (key === "debug.install") {
          return "install";
        }

        return "empty";
      });
      const ui = {
        showMessage(
          level: "info" | "warn" | "error",
          message: string,
          modal: boolean,
          ...items: string[]
        ): Promise<Result<string | undefined, FxError>> {
          return Promise.resolve(ok("install"));
        },
      } as UserInteraction;
      certManager = new LocalCertificateManager(ui);
      const userConfirm = await (certManager as any).waitForUserConfirm();
      chai.assert.isTrue(userConfirm);
    });

    it("waitForUserConfirm twice", async () => {
      vi.spyOn(localCertificateManagerDeps, "localize").mockImplementation((key, ...params) => {
        if (key === "debug.install") {
          return "install";
        } else if (key === "core.provision.learnMore") {
          return "learnmore";
        }

        return "empty";
      });
      let count = 0;
      const ui = {
        openUrl(link: string): Promise<Result<boolean, FxError>> {
          return Promise.resolve(ok(true));
        },
        showMessage(
          level: "info" | "warn" | "error",
          message: string,
          modal: boolean,
          ...items: string[]
        ): Promise<Result<string | undefined, FxError>> {
          count++;
          return Promise.resolve(ok(count > 1 ? "install" : "learnmore"));
        },
      } as UserInteraction;
      certManager = new LocalCertificateManager(ui);
      const userConfirm = await (certManager as any).waitForUserConfirm();
      chai.assert.isTrue(userConfirm);
    });

    it("trustCertificateWindows", async () => {
      vi.spyOn(localCertificateManagerDeps, "execPowerShell").mockImplementation(
        async (command: string) => {
          if (command.startsWith("(Get-ChildItem")) {
            // Command: `(Get-ChildItem -Path Cert:\\CurrentUser\\Root\\${thumbprint}).FriendlyName='${friendlyName}'`
            return "friendlyname";
          } else if (command.startsWith("Import-Certificate")) {
            // Command: `Import-Certificate -FilePath '${localCert.certPath}' -CertStoreLocation Cert:\\CurrentUser\\Root`
            return "import";
          } else {
            return "";
          }
        }
      );
      const certManager = new LocalCertificateManager();
      await (certManager as any).trustCertificateWindows(
        {
          certPath: "certPath",
          keyPath: "keyPath",
        },
        "thumbprint",
        "friendlyname"
      );
    });

    it("trustCertificate error", async () => {
      vi.spyOn(localCertificateManagerDeps, "osType").mockReturnValue("Windows_NT");
      const certManager = new LocalCertificateManager();
      (certManager as any).waitForUserConfirm = function (): Promise<boolean> {
        return Promise.reject(new Error("test"));
      };
      const cert = {
        certPath: "certPath",
        keyPath: "keyPath",
      } as any;
      await (certManager as any).trustCertificate(cert, "thumbprint", "friendlyname");
      chai.assert.isFalse(cert.isTrusted);
      chai.assert.isDefined(cert.error);
    });

    it("generateCertificate preserves the localhost certificate and RSA PEM contract", async () => {
      const certFile = path.resolve(fakeHomeDir, "localhost.crt");
      const keyFile = path.resolve(fakeHomeDir, "localhost.key");
      const writeFileStub = vi.spyOn(localCertificateManagerDeps, "writeFile").mockResolvedValue();

      const certManager = new LocalCertificateManager();
      const thumbprint = await certManager.generateCertificate(certFile, keyFile);

      chai.assert.equal(writeFileStub.mock.calls.length, 2);
      chai.assert.equal(writeFileStub.mock.calls[0][0], certFile);
      chai.assert.equal(writeFileStub.mock.calls[1][0], keyFile);
      const certContent = writeFileStub.mock.calls[0][1];
      const keyContent = writeFileStub.mock.calls[1][1];
      chai.assert.isString(certContent);
      chai.assert.isString(keyContent);
      if (typeof certContent !== "string" || typeof keyContent !== "string") {
        throw new Error("Expected certificate and private key PEM strings");
      }
      const cert = new X509Certificate(certContent);
      const nativeCert = new NodeCertificate(certContent);
      const privateKey = createPrivateKey(keyContent);
      chai.assert.equal(privateKey.asymmetricKeyDetails?.modulusLength, 4096);
      chai.assert.match(keyContent, /^-----BEGIN RSA PRIVATE KEY-----/);
      chai.assert.isTrue(nativeCert.checkPrivateKey(privateKey));
      chai.assert.isTrue(nativeCert.verify(nativeCert.publicKey));
      chai.assert.equal(cert.signatureAlgorithm.name, "RSASSA-PKCS1-v1_5");
      chai.assert.deepEqual(cert.signatureAlgorithm.hash, { name: "SHA-256" });
      chai.assert.deepEqual(cert.subjectName.getField("CN"), ["localhost"]);
      chai.assert.deepEqual(cert.issuerName.getField("CN"), ["localhost"]);
      const expiry = new Date(cert.notBefore);
      expiry.setFullYear(expiry.getFullYear() + 1);
      chai.assert.equal(cert.notAfter.getTime(), expiry.getTime());
      chai.assert.isFalse(cert.getExtension(BasicConstraintsExtension)?.ca);
      chai.assert.include(
        cert.getExtension(ExtendedKeyUsageExtension)?.usages,
        ExtendedKeyUsage.serverAuth
      );
      chai.assert.isTrue(
        cert
          .getExtension(SubjectAlternativeNameExtension)
          ?.names.items.some((name) => name.type === "dns" && name.value === "localhost")
      );
      chai.assert.equal(thumbprint, getCertThumbprint(certContent));
      chai.assert.deepEqual(certManager.verifyCertificateContent(certContent, keyContent), [
        thumbprint,
        true,
      ]);

      const server = createServer({ cert: certContent, key: keyContent }, (_request, response) => {
        response.end("local tab");
      });
      server.listen(0, "127.0.0.1");
      try {
        await once(server, "listening");
        const address = server.address();
        if (!address || typeof address === "string") {
          throw new Error("Expected a local TCP listener");
        }
        const body = await new Promise<string>((resolve, reject) => {
          const request = get(
            {
              hostname: "localhost",
              port: address.port,
              family: 4,
              ca: certContent,
              path: "/tabs/home",
            },
            (response) => {
              let content = "";
              response.setEncoding("utf8");
              response.on("data", (chunk: string) => (content += chunk));
              response.on("end", () => resolve(content));
              response.on("error", reject);
            }
          );
          request.on("error", reject);
          request.setTimeout(5000, () =>
            request.destroy(new Error("Local HTTPS request timed out"))
          );
        });
        chai.assert.equal(body, "local tab");
      } finally {
        const closed = once(server, "close");
        server.close();
        await closed;
      }
    });

    it("verifyCertificateInStore on Darwin should check keychain via shell", async () => {
      vi.spyOn(localCertificateManagerDeps, "osType").mockReturnValue("Darwin");
      vi.spyOn(localCertificateManagerDeps, "homedir").mockReturnValue(fakeHomeDir);
      const execShellStub = vi
        .spyOn(localCertificateManagerDeps, "execShell")
        .mockResolvedValue("SHA-1 hash: ABCDEF");

      const certManager = new LocalCertificateManager();
      const found = await certManager.verifyCertificateInStore("ABCDEF");

      chai.assert.isTrue(found);
      chai.assert.isTrue(execShellStub.mock.calls.length === 1);
    });

    it("trustCertificate on Darwin should run add-trusted-cert", async () => {
      vi.spyOn(localCertificateManagerDeps, "osType").mockReturnValue("Darwin");
      vi.spyOn(LocalCertificateManager.prototype as any, "waitForUserConfirm").mockResolvedValue(
        true
      );
      const execShellStub = vi
        .spyOn(localCertificateManagerDeps, "execShell")
        .mockResolvedValue("ok");

      const certManager = new LocalCertificateManager();
      const cert = {
        certPath: path.resolve(fakeHomeDir, "localhost.crt"),
        keyPath: path.resolve(fakeHomeDir, "localhost.key"),
      } as any;

      await (certManager as any).trustCertificate(cert, "thumbprint", "friendlyname");

      chai.assert.isTrue(execShellStub.mock.calls.length === 1);
      chai.assert.isTrue(cert.isTrusted);
    });
  });
});

describe("setupCertificate check only", () => {
  const sandbox = vi;
  afterEach(() => {
    vi.restoreAllMocks();
  });
  it("not found", async () => {
    vi.spyOn(localCertificateManagerDeps, "ensureDir").mockResolvedValue();
    vi.spyOn(localCertificateManagerDeps, "pathExists").mockResolvedValue(false);
    const certManager = new LocalCertificateManager();
    const res = await certManager.setupCertificate(true, true);
    chai.assert.isFalse(res.found);
  });
  it("found but not trusted", async () => {
    vi.spyOn(localCertificateManagerDeps, "ensureDir").mockResolvedValue();
    vi.spyOn(localCertificateManagerDeps, "pathExists").mockResolvedValue(true);
    vi.spyOn(localCertificateManagerDeps, "readFile").mockResolvedValue("aaa" as any);
    const certManager = new LocalCertificateManager();
    vi.spyOn(LocalCertificateManager.prototype, "verifyCertificateContent").mockReturnValue([
      "test",
      true,
    ]);
    vi.spyOn(LocalCertificateManager.prototype, "generateCertificate").mockResolvedValue("test");
    vi.spyOn(LocalCertificateManager.prototype, "verifyCertificateInStore").mockResolvedValue(
      false
    );
    const res = await certManager.setupCertificate(true, true);
    chai.assert.isTrue(res.found);
    chai.assert.isFalse(res.alreadyTrusted);
  });
});

function getCertThumbprint(certContent: string): string {
  return createHash("sha1").update(new NodeCertificate(certContent).raw).digest("hex");
}

async function createLightweightCertPair(
  options: X509CertificateCreateCommonParams = {}
): Promise<{ certPem: string; keyPem: string; thumbprint: string }> {
  const now = new Date();
  const expiry = new Date();
  expiry.setFullYear(expiry.getFullYear() + 1);

  const keys = await webcrypto.subtle.generateKey(
    {
      name: "RSASSA-PKCS1-v1_5",
      modulusLength: 1024,
      publicExponent: new Uint8Array([1, 0, 1]),
      hash: "SHA-256",
    },
    true,
    ["sign", "verify"]
  );
  const cert = await X509CertificateGenerator.create(
    {
      serialNumber: "01",
      subject: "CN=localhost",
      issuer: "CN=localhost",
      notBefore: now,
      notAfter: expiry,
      extensions: certificateExtensions(),
      ...options,
      publicKey: keys.publicKey,
      signingKey: keys.privateKey,
    },
    globalThis.crypto
  );
  const privateKey = createPrivateKey({
    key: Buffer.from(await webcrypto.subtle.exportKey("pkcs8", keys.privateKey)),
    format: "der",
    type: "pkcs8",
  });
  const certPem = cert.toString("pem");

  return {
    certPem,
    keyPem: privateKey.export({ format: "pem", type: "pkcs1" }).toString(),
    thumbprint: getCertThumbprint(certPem),
  };
}

function certificateExtensions(): Extension[] {
  return [
    new BasicConstraintsExtension(false),
    new ExtendedKeyUsageExtension([ExtendedKeyUsage.serverAuth]),
    new SubjectAlternativeNameExtension([{ type: "dns", value: "localhost" }]),
  ];
}

describe("local certificate compatibility", () => {
  afterEach(() => {
    vi.useRealTimers();
    vi.restoreAllMocks();
  });

  it("reuses a legacy Forge certificate without regenerating or trusting it again", async () => {
    vi.useFakeTimers();
    vi.setSystemTime(new Date("2026-11-25T00:00:00Z"));
    vi.spyOn(localCertificateManagerDeps, "ensureDir").mockResolvedValue();
    vi.spyOn(localCertificateManagerDeps, "pathExists").mockResolvedValue(true);
    vi.spyOn(localCertificateManagerDeps, "readFile")
      .mockResolvedValueOnce(legacyCertPem)
      .mockResolvedValueOnce(legacyKeyPem);
    const thumbprint = getCertThumbprint(legacyCertPem);
    const generate = vi.spyOn(LocalCertificateManager.prototype, "generateCertificate");
    const write = vi.spyOn(localCertificateManagerDeps, "writeFile").mockResolvedValue();
    const trustCommand = vi.spyOn(localCertificateManagerDeps, "execPowerShell");
    const storeCheck = vi
      .spyOn(LocalCertificateManager.prototype, "verifyCertificateInStore")
      .mockResolvedValue(true);
    const manager = new LocalCertificateManager();

    chai.assert.deepEqual(manager.verifyCertificateContent(legacyCertPem, legacyKeyPem), [
      thumbprint,
      true,
    ]);
    const result = await manager.setupCertificate(true);
    chai.assert.isTrue(result.found);
    chai.assert.isTrue(result.alreadyTrusted);
    chai.assert.isTrue(result.isTrusted);
    chai.assert.deepEqual(storeCheck.mock.calls, [[thumbprint]]);
    chai.assert.isFalse(generate.mock.calls.length > 0);
    chai.assert.isFalse(write.mock.calls.length > 0);
    chai.assert.isFalse(trustCommand.mock.calls.length > 0);
  });

  it("rejects a different RSA private key and malformed PEM", () => {
    const manager = new LocalCertificateManager();
    chai.assert.isFalse(
      manager.verifyCertificateContent(lightweightCertPair.certPem, legacyKeyPem)[1]
    );
    chai.assert.deepEqual(manager.verifyCertificateContent("invalid", "invalid"), [
      undefined,
      false,
    ]);
    chai.assert.deepEqual(
      manager.verifyCertificateContent(lightweightCertPair.certPem, "invalid"),
      [undefined, false]
    );
  });

  const invalidCertificates: Array<[string, () => X509CertificateCreateCommonParams]> = [
    ["wrong subject", () => ({ subject: "CN=example.com" })],
    ["wrong issuer", () => ({ issuer: "CN=example.com" })],
    ["missing common name", () => ({ subject: "O=localhost" })],
    ["not yet valid", () => ({ notBefore: new Date(Date.now() + 86400000) })],
    ["expired", () => ({ notAfter: new Date(Date.now() - 86400000) })],
    ["expires within one day", () => ({ notAfter: new Date(Date.now() + 3600000) })],
    ["missing extensions", () => ({ extensions: [] })],
    [
      "CA certificate",
      () => ({
        extensions: [new BasicConstraintsExtension(true), ...certificateExtensions().slice(1)],
      }),
    ],
    [
      "missing server-auth usage",
      () => ({
        extensions: [
          new BasicConstraintsExtension(false),
          new ExtendedKeyUsageExtension([ExtendedKeyUsage.clientAuth]),
          new SubjectAlternativeNameExtension([{ type: "dns", value: "localhost" }]),
        ],
      }),
    ],
    [
      "wrong DNS alternative name",
      () => ({
        extensions: [
          ...certificateExtensions().slice(0, 2),
          new SubjectAlternativeNameExtension([{ type: "dns", value: "example.com" }]),
        ],
      }),
    ],
    [
      "non-DNS alternative name",
      () => ({
        extensions: [
          ...certificateExtensions().slice(0, 2),
          new SubjectAlternativeNameExtension([{ type: "url", value: "localhost" }]),
        ],
      }),
    ],
  ];

  for (const [name, options] of invalidCertificates) {
    it(`rejects ${name}`, async () => {
      const pair = await createLightweightCertPair(options());
      chai.assert.deepEqual(
        new LocalCertificateManager().verifyCertificateContent(pair.certPem, pair.keyPem),
        [pair.thumbprint, false]
      );
    });
  }
});

// Generated by the previous implementation, not by the replacement library.
const legacyCertPem = `-----BEGIN CERTIFICATE-----
MIIB1jCCAT+gAwIBAgIBATANBgkqhkiG9w0BAQsFADAUMRIwEAYDVQQDEwlsb2Nh
bGhvc3QwHhcNMjAwMTAxMDAwMDAwWhcNNDAwMTAxMDAwMDAwWjAUMRIwEAYDVQQD
Ewlsb2NhbGhvc3QwgZ8wDQYJKoZIhvcNAQEBBQADgY0AMIGJAoGBAKIlR7mROADI
wzZy79AUIXcuwEUcK4BtoNEJ03ShYK49YmBsJrN295hhx7wXVvOMfqFiSKdJKfoP
7AGlIyk3ZdBGXEAfddcX5hdfe3urASiCmYqsYa0qWRBrt3+Rmi8SI+3QqqrNitKa
QzttuXLy9SLEl1XxBLcj382KujVCOTdxAgMBAAGjODA2MAkGA1UdEwQCMAAwEwYD
VR0lBAwwCgYIKwYBBQUHAwEwFAYDVR0RBA0wC4IJbG9jYWxob3N0MA0GCSqGSIb3
DQEBCwUAA4GBABKq3k8bIkTQW8G52lV2CIxDbknNRKyJ7WWcSdy6CKR3Pjc798rL
gwYoyoJxG/o+TU4P3XyZMitierPB2bCUCSFH74qywDISbyR+FMlG9Qx3be+qzvf2
ah6B73mFKKEvCToNgrJVtVTJhCX0TJW/9Y4lEj8Un6xm3p/xvyLhrTD9
-----END CERTIFICATE-----`;

const legacyKeyPem = `-----BEGIN RSA PRIVATE KEY-----
MIICWwIBAAKBgQCiJUe5kTgAyMM2cu/QFCF3LsBFHCuAbaDRCdN0oWCuPWJgbCaz
dveYYce8F1bzjH6hYkinSSn6D+wBpSMpN2XQRlxAH3XXF+YXX3t7qwEogpmKrGGt
KlkQa7d/kZovEiPt0KqqzYrSmkM7bbly8vUixJdV8QS3I9/Niro1Qjk3cQIDAQAB
AoGAZmnJncm9Z7Fmq5N5KSIt6QtcYcpjJ6EWvDHKn5X2bNdUtbZZ02zHAjZwY7sS
Lo6ExQZBAXhSzSEVXgyh4ryczyibmDhOdDaLecOanmyGP0I1kg5DPSz6GLAh+d5I
K5e6ZVjTwBL6R13WAp5GHKm4xJ0K83yRZ2a5xDISmfu9GakCQQDVbbmaYAS9YSr2
fOfI5py1Shtu8nHk4fde2hsCSiGflXXk09o7SNoL8bMtLjMNruVJjVbd3j6NysxJ
4HII315DAkEAwnzmNOSqD9AQGahdsBRc677A3RjgAw+VkSWEH67JexQriRQlkPHE
wQwglhVBpyxmPMnyXM1EM/bhfanueeKqOwJAHl/aTVUi94zPn3rRRe9EGCzu0JBi
ky9xlCc4NtA0DtEnkSRxZ4SLu61vUZvEmkrsyNX7qeHKR0W4Yz81G3wmnwJAC08+
B8dgSEkeFo/e2tVkTeUqy/CdXjAyz1zyIPGGD90tZuw//vzAJ0WwQ2n7BmCjlVdc
+56IH8l0dV1DW90dbwJAH/YKsXMUul8IydIslK3mxUA0OAelg0xrH31Ku0lYNFs6
xxeqrVklH/GIEfhIyip2uOdigxjyci2oYPQ6qgt64A==
-----END RSA PRIVATE KEY-----`;
