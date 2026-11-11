// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import axios from "axios";
import { expect, vi } from "vitest";
import { GraphSkillFileStore } from "../../../../src/component/m365/skills/graphSkillFileStore";
import type { RemoteSkillItem } from "../../../../src/component/m365/skills/skillFileSync";

const endpoint = "https://graph.microsoft.com";
const folder = { id: "folder", name: "nested", eTag: '"folder-v1"', folder: {} };
const file = {
  id: "file",
  name: "read me.txt",
  eTag: '"file-v1"',
  size: 3,
  file: { hashes: {} },
};

describe("Graph skill file store", () => {
  const request = {
    get: vi.fn(),
    post: vi.fn(),
    put: vi.fn(),
    delete: vi.fn(),
  };

  beforeEach(() => {
    vi.spyOn(axios, "create").mockReturnValue(Object.assign(axios.create(), request));
  });

  afterEach(() => {
    vi.restoreAllMocks();
    Object.values(request).forEach((mock) => mock.mockReset());
  });

  it("AC-13: lists every page using only metadata", async () => {
    request.get
      .mockResolvedValueOnce({
        status: 200,
        data: {
          value: [folder],
          "@odata.nextLink": `${endpoint}/v1.0/drives/drive/items/root/children?$skiptoken=next`,
        },
      })
      .mockResolvedValueOnce({ status: 200, data: { value: [file] } });
    const store = new GraphSkillFileStore(endpoint, "token", "drive");
    const result = await store.listChildren("root");
    expect(result.isOk() && result.value.map((item) => item.kind)).toEqual(["folder", "file"]);
    expect(request.get).toHaveBeenCalledTimes(2);
    expect(request.get.mock.calls.every(([url]) => !url.includes("/content"))).toBe(true);
  });

  it.each([
    "https://attacker.example/v1.0/drives/drive/items/root/children",
    `${endpoint}/v1.0/drives/another/items/root/children`,
    `${endpoint}/v1.0/drives/drive/items/file/content`,
    `${endpoint}/v1.0/drives/drive/items/root/children#fragment`,
  ])("AC-13: rejects unsafe pagination %s without fetching it", async (nextLink) => {
    request.get.mockResolvedValueOnce({
      status: 200,
      data: { value: [], "@odata.nextLink": nextLink },
    });
    const result = await new GraphSkillFileStore(endpoint, "token", "drive").listChildren("root");
    expect(result.isErr()).toBe(true);
    expect(request.get).toHaveBeenCalledTimes(1);
  });

  it("AC-13: rejects cyclic pagination", async () => {
    request.get.mockResolvedValue({
      status: 200,
      data: {
        value: [],
        "@odata.nextLink": `${endpoint}/v1.0/drives/drive/items/root/children`,
      },
    });
    const result = await new GraphSkillFileStore(endpoint, "token", "drive").listChildren("root");
    expect(result.isErr()).toBe(true);
    expect(request.get.mock.calls.length).toBeLessThanOrEqual(2);
  });

  it.each([{}, { value: null }, { value: [{ ...file, eTag: undefined }] }])(
    "AC-06: malformed metadata is not treated as an empty folder",
    async (data) => {
      request.get.mockResolvedValueOnce({ status: 200, data });
      const result = await new GraphSkillFileStore(endpoint, "token", "drive").listChildren("root");
      expect(result.isErr()).toBe(true);
    }
  );

  it("AC-03: sends local bytes and encodes file names without downloading content", async () => {
    request.put.mockResolvedValueOnce({ status: 201, data: file });
    const bytes = Buffer.from([0, 1, 255]);
    const result = await new GraphSkillFileStore(endpoint, "token", "drive").uploadFile(
      "root",
      file.name,
      bytes
    );
    expect(result.isOk()).toBe(true);
    expect(request.put.mock.calls[0][0]).toBe("/drives/drive/items/root:/read%20me.txt:/content");
    expect(request.put.mock.calls[0][1]).toBe(bytes);
    expect(request.get).not.toHaveBeenCalled();
  });

  it("AC-17: existing uploads and deletes carry exact version preconditions", async () => {
    request.put.mockResolvedValueOnce({ status: 200, data: file });
    request.delete.mockResolvedValueOnce({ status: 204 });
    const store = new GraphSkillFileStore(endpoint, "token", "drive");
    const existing: RemoteSkillItem = {
      id: file.id,
      name: file.name,
      kind: "file",
      eTag: file.eTag,
      size: file.size,
    };
    expect((await store.uploadFile("root", file.name, Buffer.from("new"), existing)).isOk()).toBe(
      true
    );
    expect((await store.deleteItem(existing)).isOk()).toBe(true);
    expect(request.put.mock.calls[0][2].headers["If-Match"]).toBe(file.eTag);
    expect(request.delete.mock.calls[0][1].headers["If-Match"]).toBe(file.eTag);
  });

  it("AC-03: creates folders with conflict behavior fail", async () => {
    request.post.mockResolvedValueOnce({ status: 201, data: folder });
    const result = await new GraphSkillFileStore(endpoint, "token", "drive").createFolder(
      "root",
      folder.name
    );
    expect(result.isOk()).toBe(true);
    expect(request.post.mock.calls[0][1]).toEqual({
      name: folder.name,
      folder: {},
      "@microsoft.graph.conflictBehavior": "fail",
    });
  });
});
