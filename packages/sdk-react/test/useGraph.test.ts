// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @jest-environment jsdom
 */

import { renderHook, act, waitFor } from "@testing-library/react";
import { Client, GraphError } from "@microsoft/microsoft-graph-client";
import { useGraphWithCredential } from "../src/useGraph";
import { ErrorWithCode, ErrorCode, UserInfo } from "@microsoft/teamsfx";
import * as teamsfxlib from "@microsoft/teamsfx";
import "isomorphic-fetch";

jest.mock("@microsoft/teamsfx", () => {
  return {
    TeamsUserCredential: jest.fn().mockImplementation(() => {
      return {
        async login(): Promise<void> {},
        async getToken(): Promise<null> {
          return null;
        },
        async getUserInfo(): Promise<UserInfo> {
          return {} as UserInfo;
        },
      };
    }),
    ErrorCode: jest.requireActual("@microsoft/teamsfx").ErrorCode,
    ErrorWithCode: jest.requireActual("@microsoft/teamsfx").ErrorWithCode,
  };
});

describe("useGraphWithCredential() hook tests", () => {
  beforeEach(() => {});

  afterEach(() => {
    jest.resetAllMocks();
    jest.clearAllMocks();
  });

  it("call function after initialized without options", async () => {
    let graphScope: string[] | undefined;
    const { result } = renderHook(() =>
      useGraphWithCredential(
        (graph: Client, credential: teamsfxlib.TeamsUserCredential, scope: string[]) => {
          graphScope = scope;
          return Promise.resolve("graph data");
        },
      ),
    );

    expect(result.current.reload).toBeDefined();
    expect(result.current.data).toBe(undefined);
    expect(result.current.error).toBe(undefined);
    expect(result.current.loading).toBe(true);
    await waitFor(
      () => {
        expect(result.current.data).toBe("graph data");
        expect(result.current.error).toBe(undefined);
        expect(result.current.loading).toBe(false);
        expect(graphScope && graphScope[0]).toBe("User.Read");
      },
      { interval: 1 },
    );
  });

  it("call function after initialized with options", async () => {
    jest
      .spyOn(teamsfxlib, "TeamsUserCredential")
      .mockImplementation((): teamsfxlib.TeamsUserCredential => {
        return {
          async login(): Promise<void> {},
          async getToken(): Promise<null> {
            return null;
          },
          async getUserInfo(): Promise<UserInfo> {
            return { displayName: "testUser" } as UserInfo;
          },
        };
      });

    const teamsUserCredential = new teamsfxlib.TeamsUserCredential({
      clientId: "clientId",
      initiateLoginEndpoint: "initiateLoginEndpoint",
    });
    let graphScope: string[] | undefined;
    let fetchedCredential: teamsfxlib.TeamsUserCredential | undefined;
    const { result } = renderHook(() =>
      useGraphWithCredential(
        (graph: Client, credential: teamsfxlib.TeamsUserCredential, scope: string[]) => {
          graphScope = scope;
          fetchedCredential = credential;
          return Promise.resolve(fetchedCredential.getUserInfo());
        },
        {
          scope: ["User.Read.All"],
          credential: teamsUserCredential,
        },
      ),
    );

    expect(result.current.reload).toBeDefined();
    expect(result.current.data).toBe(undefined);
    expect(result.current.error).toBe(undefined);
    expect(result.current.loading).toBe(true);
    await waitFor(
      () => {
        expect(result.current.data).toStrictEqual({ displayName: "testUser" });
        expect(result.current.error).toBe(undefined);
        expect(result.current.loading).toBe(false);
        expect(graphScope && graphScope[0]).toBe("User.Read.All");
      },
      { interval: 1 },
    );
  });

  it("call login() automatically when user has not consented", async () => {
    jest
      .spyOn(teamsfxlib, "TeamsUserCredential")
      .mockImplementation((): teamsfxlib.TeamsUserCredential => {
        return {
          async login(): Promise<void> {},
          async getToken(): Promise<null> {
            return null;
          },
          async getUserInfo(): Promise<UserInfo> {
            return {} as UserInfo;
          },
        };
      });
    let graphScope: string[] | undefined;
    let callTime = 0;

    const { result } = renderHook(() =>
      useGraphWithCredential(
        (graph: Client, credential: teamsfxlib.TeamsUserCredential, scope: string[]) => {
          graphScope = scope;
          if (callTime === 0) {
            callTime++;
            const error = new GraphError();
            error.code = ErrorCode.UiRequiredError;
            return Promise.reject(error);
          } else {
            return Promise.resolve("graph data");
          }
        },
      ),
    );
    expect(result.current.reload).toBeDefined();
    expect(result.current.data).toBe(undefined);
    expect(result.current.error).toBe(undefined);
    expect(result.current.loading).toBe(true);

    await waitFor(
      () => {
        expect(result.current.data).toBe(undefined);
        expect(result.current.error).toBe(undefined);
        expect(result.current.loading).toBe(false);
        expect(graphScope && graphScope[0]).toBe("User.Read");
      },
      { interval: 1 },
    );

    act(() => result.current.reload());
    await waitFor(
      () => {
        expect(result.current.data).toBe("graph data");
      },
      { interval: 1 },
    );
  });

  it("shows error message when user cancels consent dialog", async () => {
    jest
      .spyOn(teamsfxlib, "TeamsUserCredential")
      .mockImplementation((): teamsfxlib.TeamsUserCredential => {
        return {
          async login(): Promise<void> {
            throw new ErrorWithCode("CancelledByUser");
          },
          async getToken(): Promise<null> {
            return null;
          },
          async getUserInfo(): Promise<UserInfo> {
            return {} as UserInfo;
          },
        };
      });

    let graphScope: string[] | undefined;
    const { result } = renderHook(() =>
      useGraphWithCredential(
        (graph: Client, credential: teamsfxlib.TeamsUserCredential, scope: string[]) => {
          graphScope = scope;
          const error = new GraphError();
          error.code = ErrorCode.UiRequiredError;
          return Promise.reject(error);
        },
      ),
    );

    expect(result.current.reload).toBeDefined();
    expect(result.current.data).toBe(undefined);
    expect(result.current.error).toBe(undefined);
    expect(result.current.loading).toBe(true);

    await waitFor(
      () => {
        expect(result.current.data).toBe(undefined);
        expect(result.current.error).toBe(undefined);
        expect(result.current.loading).toBe(false);

        expect(graphScope && graphScope[0]).toBe("User.Read");
      },
      { interval: 1 },
    );

    act(() => result.current.reload());
    await waitFor(
      () => {
        expect(result.current.data).toBe(undefined);
        expect(result.current.error).toBeDefined();
        expect((result.current.error as ErrorWithCode).message).toBe(
          "CancelledByUser" +
            '\nIf you see "AADSTS50011: The reply URL specified in the request does not match the reply URLs configured for the application" ' +
            "in the popup window, you may be using unmatched version for TeamsFx SDK (version >= 0.5.0) and Teams Toolkit (version < 3.3.0) or " +
            "cli (version < 0.11.0). Please refer to the help link for how to fix the issue: " +
            "https://learn.microsoft.com/microsoftteams/platform/toolkit/develop-single-sign-on-experience-in-teams",
        );
      },
      { interval: 1 },
    );
  });

  it("throws unknown error", async () => {
    jest
      .spyOn(teamsfxlib, "TeamsUserCredential")
      .mockImplementation((): teamsfxlib.TeamsUserCredential => {
        return {
          async login(): Promise<void> {
            throw new ErrorWithCode("CancelledByUser");
          },
          async getToken(): Promise<null> {
            return null;
          },
          async getUserInfo(): Promise<UserInfo> {
            return {} as UserInfo;
          },
        };
      });
    let graphScope: string[] | undefined;
    const { result } = renderHook(() =>
      useGraphWithCredential(
        (graph: Client, credential: teamsfxlib.TeamsUserCredential, scope: string[]) => {
          graphScope = scope;
          const error = new Error("unknown error");
          return Promise.reject(error);
        },
      ),
    );

    expect(result.current.reload).toBeDefined();
    expect(result.current.data).toBe(undefined);
    expect(result.current.error).toBe(undefined);
    expect(result.current.loading).toBe(true);

    await waitFor(
      () => {
        expect(result.current.data).toBe(undefined);
        expect(result.current.error).toStrictEqual(new Error("unknown error"));
        expect(result.current.loading).toBe(false);
        expect(graphScope && graphScope[0]).toBe("User.Read");
      },
      { interval: 1 },
    );
  });
});
