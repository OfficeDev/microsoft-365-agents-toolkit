// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import {
  LogLevel,
  setLogLevel,
  setLogFunction,
  TeamsUserCredentialAuthConfig,
  TeamsUserCredential,
} from "@microsoft/teamsfx";
import { useTeams } from "./useTeams";
import { Theme } from "@fluentui/react-components";
import { useData } from "./useData";

/**
 * @deprecated This package will be deprecated by 2026-07. Please use [Microsoft 365 Agents SDK](https://www.npmjs.com/package/@microsoft/agents-hosting) instead.
 */
export type TeamsContextWithCredential = {
  /**
   * Instance of TeamsUserCredential.
   */
  teamsUserCredential?: TeamsUserCredential;
  /**
   * Status of data loading.
   */
  loading: boolean;
  /**
   * Error information.
   */
  error: unknown;
  /**
   * Indicates that current environment is in Teams
   */
  inTeams?: boolean;
  /**
   * Teams theme.
   */
  theme: Theme;
  /**
   * Teams theme string.
   */
  themeString: string;
  /**
   * Teams context object.
   */
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  context?: any;
};

/**
 * Initialize TeamsFx SDK with customized configuration.
 *
 * @param authConfig - custom configuration to override default ones.
 * @returns TeamsContextWithCredential object
 *
 * @deprecated This package will be deprecated by 2026-07. Please use [Microsoft 365 Agents SDK](https://www.npmjs.com/package/@microsoft/agents-hosting) instead.
 * @public
 */
export function useTeamsUserCredential(
  authConfig: TeamsUserCredentialAuthConfig,
): TeamsContextWithCredential {
  const [result] = useTeams({});
  const { data, error, loading } = useData(() => {
    if (process.env.NODE_ENV === "development") {
      setLogLevel(LogLevel.Verbose);
      setLogFunction((level: LogLevel, message: string) => {
        console.log(message);
      });
    }
    return Promise.resolve(new TeamsUserCredential(authConfig));
  });
  return {
    ...result,
    teamsUserCredential: data,
    error,
    loading: loading || (result.loading ?? true),
  };
}
