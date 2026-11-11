// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { FxError, Result, SystemError, err, ok } from "@microsoft/teamsfx-api";
import axios from "axios";
import MockAdapter from "axios-mock-adapter";
import { createHash } from "crypto";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { GraphSkillFileStore } from "../../../../src/component/m365/skills/graphSkillFileStore";
import {
  LocalSkillFile,
  RemoteSkillItem,
  SkillFileStore,
  synchronizeSkillFiles,
} from "../../../../src/component/m365/skills/skillFileSync";

const failure = new SystemError("test", "Conflict", "conflict");

function file(path: string, content = "abc"): LocalSkillFile {
  return { path, content: Buffer.from(content) };
}

function remote(id: string, name: string, content = "abc"): RemoteSkillItem {
  return {
    id,
    name,
    kind: "file",
    eTag: `"${id}:1"`,
    size: Buffer.byteLength(content),
    hashes: { sha1Hash: createHash("sha1").update(content).digest("hex") },
  };
}

function folder(id: string, name: string): RemoteSkillItem {
  return { id, name, kind: "folder", eTag: `"${id}:1"` };
}

class MemoryStore implements SkillFileStore {
  readonly children = new Map<string, RemoteSkillItem[]>([["root", []]]);
  readonly calls: string[] = [];
  readonly uploaded = new Map<string, Buffer>();
  readonly deleted: RemoteSkillItem[] = [];
  readonly updated: RemoteSkillItem[] = [];
  failList?: string;
  failUpload?: string;
  failDelete?: string;
  failCreate?: string;
  changeFolderETagOnDelete = false;
  active = 0;
  maxActive = 0;
  private nextId = 0;

  private async operation(name: string): Promise<void> {
    this.calls.push(name);
    this.active++;
    this.maxActive = Math.max(this.maxActive, this.active);
    await Promise.resolve();
    this.active--;
  }

  async listChildren(id: string): Promise<Result<RemoteSkillItem[], FxError>> {
    await this.operation(`list:${id}`);
    const items = this.children.get(id);
    return this.failList === id ? err(failure) : ok(items === undefined ? [] : items);
  }

  async createFolder(parentId: string, name: string): Promise<Result<RemoteSkillItem, FxError>> {
    await this.operation(`create:${parentId}/${name}`);
    if (this.failCreate === name) return err(failure);
    const item = folder(`created-${this.nextId++}`, name);
    this.children.set(parentId, [...(this.children.get(parentId) ?? []), item]);
    this.children.set(item.id, []);
    return ok(item);
  }

  async uploadFile(
    parentId: string,
    name: string,
    content: Buffer,
    existing?: RemoteSkillItem
  ): Promise<Result<RemoteSkillItem, FxError>> {
    await this.operation(`upload:${parentId}/${name}`);
    if (this.failUpload === name) return err(failure);
    if (existing) this.updated.push(existing);
    const item = remote(existing?.id ?? `uploaded-${this.nextId++}`, name);
    item.size = content.length;
    item.hashes = { sha1Hash: createHash("sha1").update(content).digest("hex") };
    this.children.set(parentId, [
      ...(this.children.get(parentId) ?? []).filter((child) => child.id !== item.id),
      item,
    ]);
    this.uploaded.set(item.id, content);
    return ok(item);
  }

  async deleteItem(item: RemoteSkillItem): Promise<Result<void, FxError>> {
    await this.operation(`delete:${item.id}`);
    if (this.failDelete === item.id) return err(failure);
    const current = [...this.children.values()].flat().find((child) => child.id === item.id);
    if (this.changeFolderETagOnDelete && current?.eTag !== item.eTag) return err(failure);
    this.deleted.push(item);
    for (const [id, items] of this.children) {
      this.children.set(
        id,
        items.filter((child) => child.id !== item.id)
      );
      if (this.changeFolderETagOnDelete && items.some((child) => child.id === item.id)) {
        for (const [parentId, siblings] of this.children) {
          this.children.set(
            parentId,
            siblings.map((sibling) =>
              sibling.id === id ? { ...sibling, eTag: `${sibling.eTag}-changed` } : sibling
            )
          );
        }
      }
    }
    this.children.delete(item.id);
    return ok(undefined);
  }
}

function expectError(result: { isErr(): boolean; error?: FxError }): void {
  expect(result.isErr()).toBe(true);
}

describe("synchronizeSkillFiles", () => {
  it("AC-03: uploads nested, binary, and empty files without modifying inputs", async () => {
    const store = new MemoryStore();
    const files = [
      file("nested/deep/empty", ""),
      { path: "binary", content: Buffer.from([0, 255]) },
    ];
    const result = await synchronizeSkillFiles(store, "root", files);
    expect(result._unsafeUnwrap()).toEqual({ uploaded: 2, skipped: 0, deleted: 0 });
    expect([...store.uploaded.values()]).toEqual([Buffer.alloc(0), Buffer.from([0, 255])]);
    expect(files[1].content).toEqual(Buffer.from([0, 255]));
    expect(store.maxActive).toBe(1);
  });

  describe("synchronizeSkillFiles with the real GraphSkillFileStore", () => {
    let http: MockAdapter;
    let store: GraphSkillFileStore;
    const endpoint = "https://graph.microsoft.com";
    const rootChildren = /\/drives\/drive\/items\/root\/children(?:\?.*)?$/;
    const nestedChildren = /\/drives\/drive\/items\/nested-id\/children(?:\?.*)?$/;
    const nested = { id: "nested-id", name: "nested", eTag: '"nested-v1"', folder: {} };

    function wireFile(id: string, name: string, content: Buffer) {
      return {
        id,
        name,
        eTag: `"${id}-v1"`,
        size: content.length,
        file: { hashes: { sha1Hash: createHash("sha1").update(content).digest("hex") } },
      };
    }

    beforeEach(() => {
      const client = axios.create();
      // Unmatched requests fail locally; no test can send real network traffic.
      http = new MockAdapter(client, { onNoMatch: "throwException" });
      vi.spyOn(axios, "create").mockReturnValue(client);
      store = new GraphSkillFileStore(endpoint, "test-token", "drive");
    });

    afterEach(() => {
      http.restore();
      vi.restoreAllMocks();
    });

    it("AC-03/AC-04: nested binary/empty uploads converge to a metadata-only hash no-op", async () => {
      const binary = Buffer.from([0, 255, 128]);
      const empty = Buffer.alloc(0);
      const binaryItem = wireFile("binary-id", "binary file.bin", binary);
      const emptyItem = wireFile("empty-id", "empty", empty);
      http.onGet(rootChildren).replyOnce(200, { value: [] });
      http.onPost(rootChildren).replyOnce(201, nested);
      http
        .onPut("/drives/drive/items/nested-id:/binary%20file.bin:/content")
        .replyOnce(201, binaryItem);
      http.onPut("/drives/drive/items/nested-id:/empty:/content").replyOnce(201, emptyItem);
      const files = [
        { path: "nested/binary file.bin", content: binary },
        { path: "nested/empty", content: empty },
      ];

      expect((await synchronizeSkillFiles(store, "root", files))._unsafeUnwrap()).toEqual({
        uploaded: 2,
        skipped: 0,
        deleted: 0,
      });
      expect(JSON.parse(http.history.post[0].data)).toEqual({
        name: "nested",
        folder: {},
        "@microsoft.graph.conflictBehavior": "fail",
      });
      expect(http.history.put.map((request) => request.data)).toEqual([binary, empty]);

      http.resetHistory();
      http.onGet(rootChildren).replyOnce(200, { value: [nested] });
      // QuickXor metadata also traverses the real wire parser, not only a fake port.
      http.onGet(nestedChildren).replyOnce(200, {
        value: [
          binaryItem,
          { ...emptyItem, file: { hashes: { quickXorHash: "AAAAAAAAAAAAAAAAAAAAAAAAAAA=" } } },
        ],
      });
      expect((await synchronizeSkillFiles(store, "root", files))._unsafeUnwrap()).toEqual({
        uploaded: 0,
        skipped: 2,
        deleted: 0,
      });
      expect(http.history.get).toHaveLength(2);
      expect(http.history.get.every((request) => !request.url?.includes("/content"))).toBe(true);
      expect(http.history.put).toHaveLength(0);
      expect(http.history.post).toHaveLength(0);
      expect(http.history.delete).toHaveLength(0);
    });

    it("AC-09/AC-17: midway HTTP conflict stops subsequent uploads without retry", async () => {
      const original = wireFile("existing", "b", Buffer.from("old"));
      let conflicts = 0;
      http.onGet(rootChildren).replyOnce(200, { value: [original] });
      http
        .onPut("/drives/drive/items/root:/a:/content")
        .replyOnce(201, wireFile("a-id", "a", Buffer.from("abc")));
      http.onPut(/\/drives\/drive\/items\/existing\/content(?:\?.*)?$/).reply(() => {
        conflicts++;
        return [412, { error: { code: "preconditionFailed", message: "Version changed" } }];
      });

      expect(
        (
          await synchronizeSkillFiles(store, "root", [file("a"), file("b", "new"), file("c")])
        ).isErr()
      ).toBe(true);
      expect(http.history.put).toHaveLength(2);
      expect(conflicts).toBe(1);
      expect(http.history.put[1].url).toBe("/drives/drive/items/existing/content");
      expect(http.history.put[1].headers?.["If-Match"]).toBe(original.eTag);
      expect(http.history.put[1].data).toEqual(Buffer.from("new"));
      expect(http.history.post).toHaveLength(0);
      expect(http.history.delete).toHaveLength(0);
    });

    it("AC-17: folder refresh and emptiness checks precede deletion with the new eTag", async () => {
      const child = wireFile("child-id", "a", Buffer.from("abc"));
      const refreshed = { ...nested, eTag: '"nested-after-child-delete"' };
      const sequence: string[] = [];
      http.onGet(rootChildren).replyOnce(() => {
        sequence.push("inventory-root");
        return [200, { value: [nested] }];
      });
      http.onGet(nestedChildren).replyOnce(() => {
        sequence.push("inventory-folder");
        return [200, { value: [child] }];
      });
      http.onDelete("/drives/drive/items/child-id").replyOnce((request) => {
        sequence.push("delete-child");
        expect(request.headers?.["If-Match"]).toBe(child.eTag);
        return [204];
      });
      http.onGet(rootChildren).replyOnce(() => {
        sequence.push("refresh-folder");
        return [200, { value: [refreshed] }];
      });
      http.onGet(nestedChildren).replyOnce(() => {
        sequence.push("check-empty");
        return [200, { value: [] }];
      });
      http.onDelete("/drives/drive/items/nested-id").replyOnce((request) => {
        sequence.push("delete-folder");
        expect(request.headers?.["If-Match"]).toBe(refreshed.eTag);
        return [204];
      });

      expect((await synchronizeSkillFiles(store, "root", []))._unsafeUnwrap()).toEqual({
        uploaded: 0,
        skipped: 0,
        deleted: 2,
      });
      expect(sequence).toEqual([
        "inventory-root",
        "inventory-folder",
        "delete-child",
        "refresh-folder",
        "check-empty",
        "delete-folder",
      ]);
      expect(http.history.delete.map((request) => request.url)).toEqual([
        "/drives/drive/items/child-id",
        "/drives/drive/items/nested-id",
      ]);
    });
  });

  it.each([
    ["", "da39a3ee5e6b4b0d3255bfef95601890afd80709"],
    ["abc", "a9993e364706816aba3e25717850c26c9cd0d89d"],
    [
      "abcdbcdecdefdefgefghfghighijhijkijkljklmklmnlmnomnopnopq",
      "84983e441c3bd26ebaae4aa1f95129e5e54670f1",
    ],
  ])(
    "AC-04: verifies independent SHA-1 vector %s and supports uppercase hashes",
    async (content, hash) => {
      const store = new MemoryStore();
      store.children.set("root", [
        { ...remote("a", "a", content), hashes: { sha1Hash: hash.toUpperCase() } },
      ]);
      expect(
        (await synchronizeSkillFiles(store, "root", [file("a", content)]))._unsafeUnwrap()
      ).toEqual({ uploaded: 0, skipped: 1, deleted: 0 });
      expect(store.calls).toEqual(["list:root"]);
    }
  );

  it("AC-04: skips unchanged files on retry but detects same-size edits and size mismatches", async () => {
    const store = new MemoryStore();
    await synchronizeSkillFiles(store, "root", [file("a")]);
    expect((await synchronizeSkillFiles(store, "root", [file("a")]))._unsafeUnwrap().skipped).toBe(
      1
    );
    const prior = store.children.get("root")?.[0];
    expect(
      (await synchronizeSkillFiles(store, "root", [file("a", "xyz")]))._unsafeUnwrap().uploaded
    ).toBe(1);
    expect(store.updated[0]).toBe(prior);
    const current = store.children.get("root")?.[0];
    if (current) current.size = 4;
    expect(
      (await synchronizeSkillFiles(store, "root", [file("a", "xyz")]))._unsafeUnwrap().uploaded
    ).toBe(1);
  });

  // Independent binary vectors from rclone's QuickXorHash tests:
  // https://github.com/rclone/rclone/blob/master/backend/onedrive/quickxorhash/quickxorhash_test.go
  /* eslint-disable no-secrets/no-secrets -- Public hash test vectors, not credentials. */
  it.each([
    ["", "AAAAAAAAAAAAAAAAAAAAAAAAAAA="],
    ["Sg==", "SgAAAAAAAAAAAAAAAQAAAAAAAAA="],
    ["tbQ=", "taAFAAAAAAAAAAAAAgAAAAAAAAA="],
    ["0pZP", "0rDEEwAAAAAAAAAAAwAAAAAAAAA="],
    ["gQaybEqS/4UlDc8e4IJm", "gDCALNigBEn8oxAlZ8AzPAAOQZg="],
    ["2fuxhBJXtpWFe8dOfdGeHw==", "O9tHLAghgSvYohKFyMMxnNCHaHg="],
    [
      "qwGf2ESubE5jOUHHyc94ORczFYYbc2OmEzo+hBIyzJiNwAzC8PvJqtTzwkWkSslgHFGWQZR2BV5+" +
        "uYTrYT7HVwRM40vqfj0dBgeDENyTenIOL1LHkjtDKoXEnQ0mXAHoJ8PjbNC93zi5TovVRXTNzfGE" +
        "s5dpWVqxUzb5lc7dwkyvOluBw482mQ4xrzYyIY1t+//OrNi1ObGXuUw2jBQOFfJVj2Y6BOyYmfB1" +
        "y36eBxi3zxeG5d5NYjm2GSh6e08QMAwu3zrINcqIzLOuNIiGXBtl7DjKt7b5wqi4oFiRpZsCyx2s" +
        "mhSrdrtK/CkdU6nDN+34vSR/M8rZpWQdBE7a8g==",
      "WYT9JY3JIo/pEBp+tIM6Gt2nyTM=",
    ],
    /* eslint-enable no-secrets/no-secrets */
  ])("AC-04: verifies independent QuickXor binary vector %#", async (input, quickXorHash) => {
    const store = new MemoryStore();
    const content = Buffer.from(input, "base64");
    store.children.set("root", [
      {
        ...remote("a", "a"),
        size: content.length,
        hashes: { quickXorHash },
      },
    ]);
    expect(
      (await synchronizeSkillFiles(store, "root", [{ path: "a", content }]))._unsafeUnwrap()
    ).toEqual({ uploaded: 0, skipped: 1, deleted: 0 });
    expect(store.calls).toEqual(["list:root"]);
  });

  it("AC-04: QuickXor detects same-size edits and requires matching size", async () => {
    const store = new MemoryStore();
    const item = {
      ...remote("a", "a"),
      size: 1,
      hashes: { quickXorHash: "SgAAAAAAAAAAAAAAAQAAAAAAAAA=" },
    };
    store.children.set("root", [item]);
    expect(
      (await synchronizeSkillFiles(store, "root", [file("a", "K")]))._unsafeUnwrap().uploaded
    ).toBe(1);
    store.children.set("root", [{ ...item, size: 2 }]);
    expect(
      (await synchronizeSkillFiles(store, "root", [file("a", "J")]))._unsafeUnwrap().uploaded
    ).toBe(1);
  });

  it("AC-06: a valid SHA-1 mismatch cannot be overridden by matching QuickXor", async () => {
    const store = new MemoryStore();
    store.children.set("root", [
      {
        ...remote("a", "a", "K"),
        hashes: { sha1Hash: "0".repeat(40), quickXorHash: "SgAAAAAAAAAAAAAAAQAAAAAAAAA=" },
      },
    ]);
    expect(
      (await synchronizeSkillFiles(store, "root", [file("a", "J")]))._unsafeUnwrap().uploaded
    ).toBe(1);
  });

  it.each([
    "AAAAAAAAAAAAAAAAAAAAAA==",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAA=\n",
    "AAAAAAAAAAAAAAAAAAAAAAAAAA!=",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAB=",
  ])("AC-06: malformed or noncanonical QuickXor %j always uploads", async (quickXorHash) => {
    const store = new MemoryStore();
    store.children.set("root", [{ ...remote("a", "a", ""), hashes: { quickXorHash } }]);
    expect(
      (await synchronizeSkillFiles(store, "root", [file("a", "")]))._unsafeUnwrap().uploaded
    ).toBe(1);
  });

  it.each([
    undefined,
    {},
    { sha1Hash: "" },
    { sha1Hash: "not-a-sha1" },
    { sha1Hash: "a".repeat(39) },
    { quickXorHash: "AAAAAAAAAAAAAAAAAAAAAA==" },
    { sha1Hash: 123 },
  ])("AC-06: uploads conservatively for unsupported or malformed hashes %j", async (hashes) => {
    const store = new MemoryStore();
    store.children.set("root", [JSON.parse(JSON.stringify({ ...remote("a", "a"), hashes }))]);
    expect((await synchronizeSkillFiles(store, "root", [file("a")]))._unsafeUnwrap().uploaded).toBe(
      1
    );
  });

  it("AC-06: uploads conservatively when size is unavailable", async () => {
    const store = new MemoryStore();
    const item = remote("a", "a");
    delete item.size;
    store.children.set("root", [item]);
    expect((await synchronizeSkillFiles(store, "root", [file("a")]))._unsafeUnwrap().uploaded).toBe(
      1
    );
  });

  it("AC-05: add/edit/delete/rename and both file-folder replacements converge bottom-up", async () => {
    const store = new MemoryStore();
    const staleFolder = folder("old", "old");
    const staleChild = remote("old-child", "child");
    store.children.set("root", [
      staleFolder,
      remote("to-folder", "nested"),
      remote("edit", "edit"),
      remote("rename", "before"),
    ]);
    store.children.set("old", [folder("old-nested", "nested")]);
    store.children.set("old-nested", [staleChild]);
    const result = await synchronizeSkillFiles(store, "root", [
      file("old", "replacement"),
      file("nested/new"),
      file("edit", "xyz"),
      file("after"),
    ]);
    expect(result._unsafeUnwrap()).toEqual({ uploaded: 4, skipped: 0, deleted: 5 });
    expect(store.calls.slice(0, 3)).toEqual(["list:root", "list:old", "list:old-nested"]);
    expect(store.deleted.indexOf(staleChild)).toBeLessThan(store.deleted.indexOf(staleFolder));
    expect(store.deleted).not.toContainEqual(expect.objectContaining({ id: "root" }));
    expect(
      (
        await synchronizeSkillFiles(store, "root", [
          file("old", "replacement"),
          file("nested/new"),
          file("edit", "xyz"),
          file("after"),
        ])
      )._unsafeUnwrap()
    ).toEqual({ uploaded: 0, skipped: 4, deleted: 0 });
  });

  it("AC-05: treats cross-inventory case changes as renames", async () => {
    const store = new MemoryStore();
    store.children.set("root", [folder("dir", "DIR")]);
    store.children.set("dir", [remote("child", "a")]);
    expect((await synchronizeSkillFiles(store, "root", [file("dir/a")]))._unsafeUnwrap()).toEqual({
      uploaded: 1,
      skipped: 0,
      deleted: 2,
    });
  });

  it("AC-05: removes an empty local inventory's descendants but never the root", async () => {
    const store = new MemoryStore();
    store.children.set("root", [folder("empty", "empty"), remote("a", "a")]);
    expect((await synchronizeSkillFiles(store, "root", []))._unsafeUnwrap()).toEqual({
      uploaded: 0,
      skipped: 0,
      deleted: 2,
    });
    expect(store.children.has("root")).toBe(true);
  });

  it("AC-06: complete nested inventory failure prevents every mutation", async () => {
    const store = new MemoryStore();
    store.children.set("root", [remote("stale", "stale"), folder("dir", "dir")]);
    store.failList = "dir";
    const result = await synchronizeSkillFiles(store, "root", [file("new")]);
    expect(result._unsafeUnwrapErr()).toBe(failure);
    expect(store.calls).toEqual(["list:root", "list:dir"]);
  });

  it("AC-09: a failed upload retains progress and retry re-reads current metadata", async () => {
    const store = new MemoryStore();
    store.failUpload = "b";
    expect(
      (
        await synchronizeSkillFiles(store, "root", [file("a"), file("b"), file("c")])
      )._unsafeUnwrapErr()
    ).toBe(failure);
    expect(store.calls).toEqual(["list:root", "upload:root/a", "upload:root/b"]);
    store.failUpload = undefined;
    expect(
      (
        await synchronizeSkillFiles(store, "root", [file("a"), file("b"), file("c")])
      )._unsafeUnwrap()
    ).toEqual({ uploaded: 2, skipped: 1, deleted: 0 });
  });

  it.each(["", "/", "/a", "../a", "a/../b", "a//b", "a\\b", "a/", "a:b", "a/\0b", "a."])(
    "AC-12: rejects unsafe local path %j before mutation",
    async (path) => {
      const store = new MemoryStore();
      expectError(await synchronizeSkillFiles(store, "root", [file(path)]));
      expect(store.calls).toEqual([]);
    }
  );

  it.each([
    ["a", "A"],
    ["a", "a"],
    ["a", "a/b"],
    ["a/b", "A/c"],
    ["é/a", "e\u0301/b"],
  ])("AC-12: rejects local duplicate/case/type collisions %j %j", async (a, b) => {
    const store = new MemoryStore();
    expectError(await synchronizeSkillFiles(store, "root", [file(a), file(b)]));
    expect(store.calls).toEqual([]);
  });

  it.each(
    [
      [remote("a", "../a")],
      [remote("a", "a"), remote("b", "A")],
      [remote("a", "a"), remote("a", "b")],
      [remote("root", "a")],
      [{ ...remote("a", "a"), id: "" }],
      [{ ...remote("a", "a"), eTag: "" }],
      [{ ...remote("a", "a"), kind: "link" }],
      [{ ...remote("a", "a"), size: -1 }],
      [null],
      null,
    ].map((items) => ({ items }))
  )("AC-12: rejects malformed remote inventory $items before mutation", async ({ items }) => {
    const store = new MemoryStore();
    store.children.set("root", JSON.parse(JSON.stringify(items)));
    expectError(await synchronizeSkillFiles(store, "root", [file("new")]));
    expect(store.calls).toEqual(["list:root"]);
  });

  it("AC-12: rejects repeated IDs across nested folders before deleting anything", async () => {
    const store = new MemoryStore();
    store.children.set("root", [folder("dir", "dir"), remote("repeated", "a")]);
    store.children.set("dir", [remote("repeated", "b")]);
    expectError(await synchronizeSkillFiles(store, "root", []));
    expect(store.calls).toEqual(["list:root", "list:dir"]);
  });

  it("AC-12: rejects a folder cycle to the root before deleting anything", async () => {
    const store = new MemoryStore();
    store.children.set("root", [folder("dir", "dir")]);
    store.children.set("dir", [folder("root", "cycle")]);
    expectError(await synchronizeSkillFiles(store, "root", []));
    expect(store.calls).toEqual(["list:root", "list:dir"]);
  });

  it("AC-12: refuses to upload through a newly created folder aliasing the root", async () => {
    const store = new MemoryStore();
    vi.spyOn(store, "createFolder").mockResolvedValue(ok(folder("root", "dir")));
    expectError(await synchronizeSkillFiles(store, "root", [file("dir/a")]));
    expect(store.calls).toEqual(["list:root"]);
  });

  it("AC-12: rejects a malformed upload result and stops subsequent uploads", async () => {
    const store = new MemoryStore();
    const upload = vi.spyOn(store, "uploadFile").mockResolvedValue(ok(remote("root", "a")));
    expectError(await synchronizeSkillFiles(store, "root", [file("a"), file("b")]));
    expect(upload).toHaveBeenCalledTimes(1);
  });

  it("AC-12: rejects an invalid root before remote reads", async () => {
    const store = new MemoryStore();
    expectError(await synchronizeSkillFiles(store, "", []));
    expect(store.calls).toEqual([]);
  });

  it("AC-17: forwards exact versioned items and stops on deletion conflicts without retries", async () => {
    const store = new MemoryStore();
    const child = remote("child", "child");
    store.children.set("root", [folder("dir", "dir"), remote("after", "after")]);
    store.children.set("dir", [child]);
    store.failDelete = "child";
    const deletion = vi.spyOn(store, "deleteItem");
    expect((await synchronizeSkillFiles(store, "root", []))._unsafeUnwrapErr()).toBe(failure);
    expect(store.calls).toEqual(["list:root", "list:dir", "delete:child"]);
    expect(deletion).toHaveBeenCalledExactlyOnceWith(child);
    expect(store.deleted).toEqual([]);
  });

  it("AC-17: refreshes changed folder eTags after deleting children, then confirms emptiness", async () => {
    const store = new MemoryStore();
    store.changeFolderETagOnDelete = true;
    store.children.set("root", [folder("outer", "outer")]);
    store.children.set("outer", [folder("inner", "inner")]);
    store.children.set("inner", [remote("a", "a")]);
    expect((await synchronizeSkillFiles(store, "root", []))._unsafeUnwrap().deleted).toBe(3);
    expect(store.calls).toEqual([
      "list:root",
      "list:outer",
      "list:inner",
      "delete:a",
      "list:outer",
      "list:inner",
      "delete:inner",
      "list:root",
      "list:outer",
      "delete:outer",
    ]);
    expect(store.deleted.slice(1).map((item) => item.eTag)).toEqual([
      '"inner:1"-changed',
      '"outer:1"-changed',
    ]);
  });

  it("AC-17: unexpected new folder contents block deletion even with a refreshed eTag", async () => {
    const store = new MemoryStore();
    store.children.set("root", [folder("dir", "dir")]);
    store.children.set("dir", []);
    const list = store.listChildren.bind(store);
    let reads = 0;
    vi.spyOn(store, "listChildren").mockImplementation(async (id) => {
      if (id === "dir" && ++reads === 2) store.children.set("dir", [remote("new", "new")]);
      return list(id);
    });
    expectError(await synchronizeSkillFiles(store, "root", []));
    expect(store.deleted).toEqual([]);
    expect(store.calls).toEqual(["list:root", "list:dir", "list:root", "list:dir"]);
  });

  it("AC-17: a replaced folder identity in the refresh blocks deletion", async () => {
    const store = new MemoryStore();
    store.children.set("root", [folder("dir", "dir")]);
    const list = store.listChildren.bind(store);
    let reads = 0;
    vi.spyOn(store, "listChildren").mockImplementation(async (id) => {
      if (id === "root" && ++reads === 2)
        store.children.set("root", [folder("replacement", "dir")]);
      return list(id);
    });
    expectError(await synchronizeSkillFiles(store, "root", []));
    expect(store.deleted).toEqual([]);
  });

  it("AC-17: a child added after the empty check causes a version conflict without retries", async () => {
    const store = new MemoryStore();
    store.changeFolderETagOnDelete = true;
    store.children.set("root", [folder("dir", "dir")]);
    const list = store.listChildren.bind(store);
    let reads = 0;
    vi.spyOn(store, "listChildren").mockImplementation(async (id) => {
      const result = await list(id);
      if (id === "dir" && ++reads === 2) {
        store.children.set("dir", [remote("new", "new")]);
        store.children.set("root", [{ ...folder("dir", "dir"), eTag: '"new-version"' }]);
      }
      return result;
    });
    expect((await synchronizeSkillFiles(store, "root", []))._unsafeUnwrapErr()).toBe(failure);
    expect(store.deleted).toEqual([]);
    expect(store.calls.filter((call) => call.startsWith("delete:"))).toEqual(["delete:dir"]);
  });

  it("AC-17: failed emptiness checks never become empty folder fallbacks", async () => {
    const store = new MemoryStore();
    store.children.set("root", [folder("dir", "dir")]);
    const list = store.listChildren.bind(store);
    let reads = 0;
    vi.spyOn(store, "listChildren").mockImplementation(async (id) => {
      if (id === "dir" && ++reads === 2) return err(failure);
      return list(id);
    });
    expect((await synchronizeSkillFiles(store, "root", []))._unsafeUnwrapErr()).toBe(failure);
    expect(store.deleted).toEqual([]);
  });

  it("AC-17: errors during folder metadata refresh stop deletion", async () => {
    const store = new MemoryStore();
    store.children.set("root", [folder("dir", "dir")]);
    const list = store.listChildren.bind(store);
    let reads = 0;
    vi.spyOn(store, "listChildren").mockImplementation(async (id) => {
      if (id === "root" && ++reads === 2) return err(failure);
      return list(id);
    });
    expect((await synchronizeSkillFiles(store, "root", []))._unsafeUnwrapErr()).toBe(failure);
    expect(store.deleted).toEqual([]);
  });

  it("AC-17: stops after folder creation failure without uploading", async () => {
    const store = new MemoryStore();
    store.failCreate = "dir";
    expect((await synchronizeSkillFiles(store, "root", [file("dir/a")]))._unsafeUnwrapErr()).toBe(
      failure
    );
    expect(store.calls).toEqual(["list:root", "create:root/dir"]);
  });

  it("AC-18: bounds active work to one operation across many files", async () => {
    const store = new MemoryStore();
    const files = Array.from({ length: 1500 }, (_, i) => file(`dir-${i % 10}/file-${i}`));
    expect((await synchronizeSkillFiles(store, "root", files))._unsafeUnwrap()).toEqual({
      uploaded: 1500,
      skipped: 0,
      deleted: 0,
    });
    expect(store.maxActive).toBe(1);
    expect((await synchronizeSkillFiles(store, "root", files))._unsafeUnwrap().skipped).toBe(1500);
    expect(store.maxActive).toBe(1);
  });
});
