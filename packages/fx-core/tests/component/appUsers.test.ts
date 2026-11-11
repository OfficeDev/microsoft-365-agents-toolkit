// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { afterEach, expect, it, vi } from "vitest";
import { GraphClient } from "../../src/client/graphClient";
import { teamsDevPortalClient } from "../../src/client/teamsDevPortalClientProvider";
import { getAppUsers } from "../../src/component/utils/appUsers";
import { MockedM365Provider } from "../core/utils";

afterEach(() => vi.restoreAllMocks());

it("AC-09: an absent native user list produces no collaborators or Graph lookups", async () => {
  vi.spyOn(teamsDevPortalClient, "getUserList").mockResolvedValue(undefined);
  const graph = vi.spyOn(GraphClient.prototype, "getUserInfoFromId");
  expect((await getAppUsers("token", "id", new MockedM365Provider()))._unsafeUnwrap()).toEqual([]);
  expect(graph).not.toHaveBeenCalled();
});

it("AC-09: enriches native access control without inventing display identities", async () => {
  vi.spyOn(teamsDevPortalClient, "getUserList").mockResolvedValue([
    { aadId: "user-id", tenantId: "tenant-id", role: "Owner" },
  ]);
  const graph = vi.spyOn(GraphClient.prototype, "getUserInfoFromId").mockResolvedValue({
    id: "user-id",
    displayName: "Owner",
    userPrincipalName: "owner@example.com",
    mail: "owner@example.com",
  });
  const result = await getAppUsers("token", "resource-id", new MockedM365Provider());
  expect(result._unsafeUnwrap()).toEqual([
    {
      aadId: "user-id",
      tenantId: "tenant-id",
      role: "Owner",
      isAdministrator: true,
      displayName: "Owner",
      userPrincipalName: "owner@example.com",
    },
  ]);
  expect(graph).toHaveBeenCalledWith("user-id");
});

it("AC-09: preserves legacy user records without extra Graph calls", async () => {
  const users = [
    {
      aadId: "user-id",
      tenantId: "tenant-id",
      isAdministrator: true,
      displayName: "Owner",
      userPrincipalName: "owner@example.com",
    },
  ];
  vi.spyOn(teamsDevPortalClient, "getUserList").mockResolvedValue(users);
  const graph = vi.spyOn(GraphClient.prototype, "getUserInfoFromId");
  expect((await getAppUsers("token", "id", new MockedM365Provider()))._unsafeUnwrap()).toEqual(
    users
  );
  expect(graph).not.toHaveBeenCalled();
});

it("AC-09: reports missing native user identity instead of returning blank names", async () => {
  vi.spyOn(teamsDevPortalClient, "getUserList").mockResolvedValue([
    { aadId: "user-id", tenantId: "tenant-id", role: "Owner" },
  ]);
  vi.spyOn(GraphClient.prototype, "getUserInfoFromId").mockResolvedValue(undefined);
  expect((await getAppUsers("token", "id", new MockedM365Provider())).isErr()).toBe(true);
});
