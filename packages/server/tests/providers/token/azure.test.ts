// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import "mocha";
import { Duplex } from "stream";
import * as chai from "chai";
import chaiAsPromised from "chai-as-promised";
import sinon from "sinon";
import ServerAzureAccountProvider from "../../../src/providers/token/azure";
import { createMessageConnection } from "vscode-jsonrpc";
import { err, ok } from "@microsoft/teamsfx-api";
import { NotImplementedError } from "@microsoft/teamsfx-core";

chai.use(chaiAsPromised);

class TestStream extends Duplex {
  _write(chunk: string, _encoding: string, done: () => void) {
    this.emit("data", chunk);
    done();
  }

  _read(_size: number) {}
}

describe("azure", () => {
  const sandbox = sinon.createSandbox();
  const up = new TestStream();
  const down = new TestStream();
  const msgConn = createMessageConnection(up as any, down as any);

  afterEach(() => {
    sandbox.restore();
  });

  it("constructor", () => {
    const azure = new ServerAzureAccountProvider(msgConn);
    chai.assert.equal(azure["connection"], msgConn);
  });

  it("getIdentityCredentialAsync", () => {
    const azure = new ServerAzureAccountProvider(msgConn);
    const res = azure.getIdentityCredentialAsync();
    res.then((data) => {
      chai.assert.isUndefined(data);
    });
  });

  it("getIdentityCredentialAsync2", async () => {
    const azure = new ServerAzureAccountProvider(msgConn);
    const promise = Promise.resolve(ok("a.eyJ1c2VySWQiOiJ0ZXN0QHRlc3QuY29tIn0=.c"));
    const stub = sandbox.stub(msgConn, "sendRequest").returns(promise);
    const identity = await azure.getIdentityCredentialAsync();
    const res = await identity?.getToken("test");
    chai.assert.isNotNull(res);
  });

  it("getIdentityCredentialAsync3", async () => {
    const azure = new ServerAzureAccountProvider(msgConn);
    const promise = Promise.resolve(ok("a.eyJ1c2VySWQiOiJ0ZXN0QHRlc3QuY29tIn0=.c"));
    const stub = sandbox.stub(msgConn, "sendRequest").returns(promise);
    const identity = await azure.getIdentityCredentialAsync();
    const res = await identity?.getToken(["test"]);
    chai.assert.isNotNull(res);
  });

  it("getIdentityCredentialAsync4", async () => {
    const azure = new ServerAzureAccountProvider(msgConn);
    const promise = Promise.resolve(err(new Error("test")));
    const stub = sandbox.stub(msgConn, "sendRequest").returns(promise);
    const identity = await azure.getIdentityCredentialAsync();
    const res = await identity?.getToken(["test"]);
    chai.assert.isNull(res);
  });

  it("getIdentityCredentialAsync returns MFA required error", async () => {
    const azure = new ServerAzureAccountProvider(msgConn);

    try {
      await azure.getIdentityCredentialAsync(false, {
        wwwAuthenticate: "test",
        scopes: [],
      });
    } catch (e) {
      chai.assert.equal((e as any).name, "MFARequiredError");
    }
  });

  it("signout", async () => {
    const azure = new ServerAzureAccountProvider(msgConn);
    chai.expect(() => azure.signout()).to.throw(NotImplementedError);
  });

  it("setStatusChangeMap", async () => {
    const azure = new ServerAzureAccountProvider(msgConn);
    chai
      .expect(() => azure.setStatusChangeMap("test", sandbox.fake()))
      .to.throw(NotImplementedError);
  });

  it("removeStatusChangeMap", async () => {
    const azure = new ServerAzureAccountProvider(msgConn);
    chai.expect(() => azure.removeStatusChangeMap("test")).to.throw(NotImplementedError);
  });

  it("switchTenant", () => {
    const azure = new ServerAzureAccountProvider(msgConn);
    chai.expect(() => azure.switchTenant("")).to.throw();
  });

  describe("getJsonObject", () => {
    const azure = new ServerAzureAccountProvider(msgConn);

    it("getJsonObject: err", async () => {
      const promise = Promise.resolve(err(new Error("test")));
      const stub = sandbox.stub(msgConn, "sendRequest").returns(promise);
      await chai.expect(azure.getJsonObject()).to.be.rejected;
      stub.restore();
    });

    it("getJsonObject: ok", () => {
      const promise = Promise.resolve(ok("test"));
      const stub = sandbox.stub(msgConn, "sendRequest").returns(promise);
      const res = azure.getJsonObject();
      res.then((data) => {
        chai.expect(data).equal("test");
      });
      stub.restore();
    });
  });

  describe("listSubscriptions", () => {
    const azure = new ServerAzureAccountProvider(msgConn);

    it("listSubscriptions: err", async () => {
      const promise = Promise.resolve(err(new Error("test")));
      const stub = sandbox.stub(msgConn, "sendRequest").returns(promise);
      await chai.expect(azure.listSubscriptions()).to.be.rejected;
      stub.restore();
    });

    it("listSubscriptions: ok", () => {
      const promise = Promise.resolve(ok("test"));
      const stub = sandbox.stub(msgConn, "sendRequest").returns(promise);
      const res = azure.listSubscriptions();
      res.then((data) => {
        chai.expect(data).equal("test");
      });
      stub.restore();
    });
  });

  describe("setSubscription", () => {
    const azure = new ServerAzureAccountProvider(msgConn);

    it("setSubscription: err", async () => {
      const promise = Promise.resolve(err(new Error("test")));
      const stub = sandbox.stub(msgConn, "sendRequest").returns(promise);
      await chai.expect(azure.setSubscription("test")).to.be.rejected;
      stub.restore();
    });

    it("setSubscription: ok", () => {
      const promise = Promise.resolve(ok("test"));
      const stub = sandbox.stub(msgConn, "sendRequest").returns(promise);
      const res = azure.setSubscription("test");
      res.then((data) => {
        chai.expect(data).equal("test");
      });
      stub.restore();
    });
  });

  it("getAccountInfo", () => {
    const azure = new ServerAzureAccountProvider(msgConn);
    chai.expect(() => azure.getAccountInfo()).to.throw();
  });

  describe("getSelectedSubscription", () => {
    const azure = new ServerAzureAccountProvider(msgConn);

    it("getSelectedSubscription: err", async () => {
      const promise = Promise.resolve(err(new Error("test")));
      const stub = sandbox.stub(msgConn, "sendRequest").returns(promise);
      await chai.expect(azure.getSelectedSubscription()).to.be.rejected;
      stub.restore();
    });

    it("getSelectedSubscription: ok", () => {
      const promise = Promise.resolve(ok("test"));
      const stub = sandbox.stub(msgConn, "sendRequest").returns(promise);
      const res = azure.getSelectedSubscription();
      res.then((data) => {
        chai.expect(data).equal("test");
      });
      stub.restore();
    });
  });
});
