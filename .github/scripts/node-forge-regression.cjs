// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
// Regression vectors adapted from digitalbazaar/forge (BSD-3-Clause).
// Copyright (c) 2010, Digital Bazaar, Inc. See ../../patches/node-forge.LICENSE.

const assert = require("node:assert/strict");
const { generateKeyPairSync, sign, verify } = require("node:crypto");
const forge = require(process.argv[2]);
const { asn1, pki, util } = forge;
const field = (type, value, constructed = false) =>
  asn1.create(asn1.Class.UNIVERSAL, type, constructed, value);

function keys(publicExponent = 65537) {
  const pair = generateKeyPairSync("rsa", {
    modulusLength: 1024,
    publicExponent,
  });
  return {
    publicKey: pki.publicKeyFromPem(
      pair.publicKey.export({ type: "spki", format: "pem" }),
    ),
    privateKey: pki.privateKeyFromPem(
      pair.privateKey.export({ type: "pkcs1", format: "pem" }),
    ),
    node: pair,
  };
}

function certificate(pair, subject, issuer, serial, extensions) {
  const cert = pki.createCertificate();
  cert.publicKey = pair.publicKey;
  cert.serialNumber = serial.toString(16).padStart(2, "0");
  cert.validity.notBefore = new Date(Date.now() - 60000);
  cert.validity.notAfter = new Date(Date.now() + 86400000);
  cert.setSubject([{ name: "commonName", value: subject }]);
  cert.setIssuer([{ name: "commonName", value: issuer }]);
  cert.setExtensions(extensions);
  cert.sign(pair.privateKey, forge.md.sha256.create());
  return pki.certificateFromPem(pki.certificateToPem(cert));
}

function rsa() {
  const pair = keys(3);
  const message = Buffer.from("Agents Toolkit RSA backport regression");
  const digest = forge.md.sha256
    .create()
    .update(message.toString())
    .digest()
    .getBytes();
  for (const withNull of [false, true]) {
    for (const extra of [false, true]) {
      const algorithm = [
        field(asn1.Type.OID, asn1.oidToDer(pki.oids.sha256).getBytes()),
      ];
      if (withNull) algorithm.push(field(asn1.Type.NULL, ""));
      if (extra)
        algorithm.push(
          field(asn1.Type.OCTETSTRING, "unexpected nested element"),
        );
      const encoded = asn1
        .toDer(
          field(
            asn1.Type.SEQUENCE,
            [
              field(asn1.Type.SEQUENCE, algorithm, true),
              field(asn1.Type.OCTETSTRING, digest),
            ],
            true,
          ),
        )
        .getBytes();
      const signature = pair.privateKey.sign(encoded, "NONE");
      if (extra) {
        assert.equal(
          verify(
            "sha256",
            message,
            pair.node.publicKey,
            Buffer.from(signature, "latin1"),
          ),
          false,
        );
        assert.throws(
          () => pair.publicKey.verify(digest, signature),
          /DigestInfo/,
        );
      } else {
        assert.equal(pair.publicKey.verify(digest, signature), true);
      }
    }
  }
  const signature = sign("sha256", message, pair.node.privateKey);
  assert.equal(
    pair.publicKey.verify(digest, signature.toString("latin1")),
    true,
  );
  const md = forge.md.sha256.create().update(message.toString());
  assert.equal(
    verify(
      "sha256",
      message,
      pair.node.publicKey,
      Buffer.from(pair.privateKey.sign(md), "latin1"),
    ),
    true,
  );
}

function pem() {
  for (const input of [
    " -----BEGIN I-----G-----END B".repeat(2000) + "\x00",
    "-----BEGIN A-----\n" + "\n".repeat(256 * 1024),
    "-----BEGIN RSA-----\n" + "\n\n".repeat(128 * 1024) + "ZZZZ",
    "-----BEGIN A-----B-----\nQUJD\n-----END A-----B-----\n",
  ]) {
    assert.throws(
      () => forge.pem.decode(input),
      /Invalid PEM formatted message/,
    );
  }
  for (const input of [
    "-----BEGIN A-----\nQUJD\n-----END A-----\n",
    "-----BEGIN A-----QUJD-----END A-----",
    "-----BEGIN A-----\n" +
      "A".repeat(512 * 1024) +
      "\n\nQUJD\n-----END A-----\n",
  ]) {
    const messages = forge.pem.decode(input);
    assert.equal(messages.length, 1);
    assert.equal(messages[0].body, "ABC");
  }
  const input =
    "-----BEGIN A-----\nProc-Type: 4,ENCRYPTED\n" +
    "DEK-Info: DES-CBC,F8143EDE5960C597\nX-Header: " +
    " ".repeat(512 * 1024) +
    "x\n\nQUJD\n-----END A-----\n";
  assert.deepEqual(forge.pem.decode(input)[0].headers, [
    { name: "X-Header", values: ["x"] },
  ]);
  assert.equal(
    forge.pem.decode("-----BEGIN A-----\n-----END A-----")[0].body,
    "",
  );
}

function pkcs12Kdf() {
  for (const iterations of [0, -1, NaN, undefined]) {
    assert.throws(
      () =>
        forge.pkcs12.generateKey(
          "test-password",
          util.createBuffer("salt"),
          3,
          iterations,
          20,
        ),
      /PKCS#12 key derivation iteration count must be >= 1/,
    );
  }
  const pair = keys();
  const info = pki.wrapRsaPrivateKey(pki.privateKeyToAsn1(pair.privateKey));
  const encrypted = pki.encryptPrivateKeyInfo(info, "test-password", {
    algorithm: "3des",
    count: 1,
  });
  assert.deepEqual(pki.decryptPrivateKeyInfo(encrypted, "test-password"), info);
  encrypted.value[0].value[1].value[1].value = "\x00";
  const cipher = forge.des.createEncryptionCipher("\x01".repeat(24));
  cipher.start("\x02".repeat(8));
  cipher.update(asn1.toDer(info));
  assert.equal(cipher.finish(), true);
  encrypted.value[1].value = cipher.output.getBytes();
  assert.throws(
    () => pki.decryptPrivateKeyInfo(encrypted, "any password"),
    /PKCS#12 key derivation iteration count must be >= 1/,
  );
}

function pkcs12Mac() {
  const pair = keys();
  const cert = certificate(pair, "fixture", "fixture", 1, []);
  const createPfx = () =>
    forge.pkcs12.toPkcs12Asn1(pair.privateKey, cert, "test-password", {
      algorithm: "3des",
      count: 1,
    });
  const valid = forge.pkcs12.pkcs12FromAsn1(
    asn1.fromDer(asn1.toDer(createPfx())),
    "test-password",
  );
  const bags = valid.getBags({ bagType: pki.oids.pkcs8ShroudedKeyBag });
  assert.equal(
    bags[pki.oids.pkcs8ShroudedKeyBag][0].key.n.toString(16),
    pair.privateKey.n.toString(16),
  );
  for (const iterations of ["\x00", ""]) {
    const pfx = createPfx();
    const certOnly = forge.pkcs12.toPkcs12Asn1(null, cert, null, {
      useMac: false,
      generateLocalKeyId: false,
    });
    pfx.value[1] = certOnly.value[1];
    const mac = forge.hmac.create();
    mac.start("sha1", "\x03".repeat(20));
    mac.update(pfx.value[1].value[1].value[0].value);
    pfx.value[2].value[0].value[1].value = mac.digest().getBytes();
    pfx.value[2].value[2].value = iterations;
    assert.throws(
      () =>
        forge.pkcs12.pkcs12FromAsn1(
          asn1.fromDer(asn1.toDer(pfx)),
          "test-password",
        ),
      /PKCS#12 MAC iteration count must be >= 1/,
    );
  }
}

function record(type, fragment) {
  return util
    .createBuffer()
    .putByte(type)
    .putByte(3)
    .putByte(1)
    .putInt16(fragment.length)
    .putBytes(fragment)
    .getBytes();
}

function connection(server, log) {
  return forge.tls.createConnection({
    server,
    verify: () => true,
    tlsDataReady: (c) => {
      log.sent += c.tlsData.getBytes();
    },
    dataReady: (c) => {
      c.data.getBytes();
    },
    heartbeatReceived: () => {
      log.heartbeats++;
      if (log.heartbeats === 1 && log.throwOnFirstHeartbeat)
        throw new Error("heartbeat callback failed");
    },
    closed: () => {},
    error: (c, error) => {
      log.errors.push(error.message);
    },
  });
}

function tlsStack() {
  const request = util.hexToBytes("1803010003010000");
  const response = util.hexToBytes("1803010003020000");
  const alert = util.hexToBytes("15030100020228");
  for (const server of [true, false]) {
    const log = { sent: "", heartbeats: 0, errors: [] };
    const c = connection(server, log);
    if (!server) c.handshake();
    assert.doesNotThrow(() =>
      c.process((server ? response : request).repeat(50000) + alert),
    );
    assert.deepEqual(log.errors, ["Handshake failure."]);
  }
  const log = { sent: "", heartbeats: 0, errors: [] };
  const c = connection(true, log);
  assert.equal(
    c.process(
      record(22, util.hexToBytes("01ffffff")) +
        record(22, "\x00").repeat(50000),
    ),
    0,
  );
  assert.equal(c.input.length(), 0);
  assert.deepEqual(log.errors, []);
  const splitLog = { sent: "", heartbeats: 0, errors: [] };
  const split = connection(true, splitLog);
  const needed = [...(response.repeat(3) + alert)].map((byte) =>
    split.process(byte),
  );
  assert.deepEqual(needed.slice(0, 8), [4, 3, 2, 1, 3, 2, 1, 0]);
  assert.deepEqual(splitLog.errors, ["Handshake failure."]);
  const callbackLog = {
    sent: "",
    heartbeats: 0,
    errors: [],
    throwOnFirstHeartbeat: true,
  };
  const callback = connection(false, callbackLog);
  callback.prepareHeartbeatRequest("");
  assert.throws(
    () => callback.process(response.repeat(2) + alert),
    /heartbeat callback failed/,
  );
  assert.equal(callback.process(), 0);
  assert.equal(callbackLog.heartbeats, 2);
  assert.deepEqual(callbackLog.errors, ["Handshake failure."]);
}

function tlsFragment() {
  const interleaved = [
    record(20, "\x01"),
    record(21, util.hexToBytes("0228")),
    record(23, "data"),
    record(24, util.hexToBytes("010000")),
  ];
  for (const server of [true, false]) {
    const first = record(
      22,
      util.hexToBytes((server ? "01" : "02") + "0000100301"),
    );
    for (const fragment of interleaved) {
      const log = { sent: "", heartbeats: 0, errors: [] };
      const c = connection(server, log);
      if (!server) c.handshake();
      log.sent = "";
      assert.doesNotThrow(() => c.process(first + fragment));
      assert.deepEqual(log.errors, ["Invalid fragmented record."]);
      assert.equal(c.fail, true);
      const alert = util.createBuffer(log.sent);
      assert.equal(alert.getByte(), 21);
      alert.getBytes(4);
      assert.equal(alert.bytes(), util.hexToBytes("020a"));
      assert.equal(c.process(fragment), 0);
      assert.deepEqual(log.errors, ["Invalid fragmented record."]);
    }
    const log = { sent: "", heartbeats: 0, errors: [] };
    const c = connection(server, log);
    if (!server) c.handshake();
    log.sent = "";
    assert.equal(c.process(first + record(22, "\x00")), 0);
    assert.deepEqual(log.errors, []);
    assert.equal(c.fail, false);
    assert.equal(log.sent, "");
  }
}

function x509() {
  const pair = keys();
  const badCertificate = (message) => ({
    message,
    error: "forge.pki.BadCertificate",
  });
  function chain(extensionsList) {
    const root = certificate(pair, "root", "root", 1, [
      { name: "basicConstraints", cA: true },
      { name: "keyUsage", keyCertSign: true },
    ]);
    const certs = [];
    let issuer = "root";
    for (const [index, extensions] of extensionsList.entries()) {
      const subject = `cert${index}`;
      certs.unshift(certificate(pair, subject, issuer, index + 2, extensions));
      issuer = subject;
    }
    certs.unshift(
      certificate(pair, "leaf", issuer, extensionsList.length + 2, []),
    );
    return pki.verifyCertificateChain(pki.createCaStore([root]), certs);
  }
  for (const id of ["2.5.29.10", "2.5.29.13"]) {
    assert.throws(
      () =>
        chain([
          [
            { id, cA: true },
            { id: "2.5.29.19", cA: false },
          ],
        ]),
      badCertificate(/basicConstraints indicates the certificate is not a CA/),
    );
    assert.throws(
      () => chain([[{ id, cA: true }]]),
      badCertificate(/missing basicConstraints extension/),
    );
    assert.throws(
      () =>
        chain([
          [
            { id, cA: true },
            { id: "2.5.29.19", cA: true, pathLenConstraint: 0 },
            { name: "keyUsage", keyCertSign: true },
          ],
          [
            { name: "basicConstraints", cA: true },
            { name: "keyUsage", keyCertSign: true },
          ],
        ]),
      badCertificate(/pathLenConstraint violated/),
    );
    assert.equal(
      chain([
        [
          { id, cA: false },
          { id: "2.5.29.19", cA: true },
          { name: "keyUsage", keyCertSign: true },
        ],
      ]),
      true,
    );
  }
  assert.equal(
    chain([
      [
        { name: "basicConstraints", cA: true, pathLenConstraint: 0 },
        { name: "keyUsage", keyCertSign: true },
      ],
    ]),
    true,
  );
}

const scenarios = {
  rsa,
  pem,
  pkcs12Kdf,
  pkcs12Mac,
  tlsStack,
  tlsFragment,
  x509,
};
assert.ok(
  Object.hasOwn(scenarios, process.argv[3]),
  "Unknown node-forge regression",
);
scenarios[process.argv[3]]();
console.log(JSON.stringify({ completed: process.argv[3] }));
