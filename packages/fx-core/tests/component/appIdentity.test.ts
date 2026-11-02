// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { err, ok, UserError } from "@microsoft/teamsfx-api";
import { FeatureFlagName, FeatureFlags } from "../../src/common/featureFlags";
import { getActiveTeamsDevPortalClient } from "../../src/client/teamsDevPortalClientProvider";
import { devPortalClient } from "../../src/client/devPortalClient";
import { legacyTeamsDevPortalClient } from "../../src/client/teamsDevPortalClient";
import { Generator } from "../../src/component/generator/generator";
import {
  getAppIdEnvName,
  defaultAppIdEnvName,
  getAppTenantId,
  resolveAppResourceId,
  resolveProjectAppResourceId,
} from "../../src/component/utils/appIdentity";
import { Validator } from "../../src/component/configManager/validator";
import { pathUtils } from "../../src/component/utils/pathUtils";
import { metadataUtil } from "../../src/component/utils/metadataUtil";
import { MockedM365Provider } from "../core/utils";

const resourceId = "11111111-1111-4111-8111-111111111111";
const manifestId = "22222222-2222-4222-8222-222222222222";

beforeEach(() => {
  vi.stubEnv(FeatureFlagName.NewDeveloperPortalApis, "true");
});

afterEach(() => {
  vi.unstubAllEnvs();
  vi.restoreAllMocks();
});

describe("DP identity compatibility", () => {
  it.each(["false", "true"])("AC-20: runtime identity selection follows flag %s", (flag) => {
    vi.stubEnv(FeatureFlagName.NewDeveloperPortalApis, flag);
    const env = { APP_ID: resourceId, TEAMS_APP_ID: manifestId, CUSTOM_ID: "" };
    expect(resolveAppResourceId(env)._unsafeUnwrap()).toBe(
      flag === "true" ? resourceId : manifestId
    );
    expect(resolveAppResourceId(env, "CUSTOM_ID")._unsafeUnwrap()).toBe(
      flag === "true" ? resourceId : ""
    );
    expect(resolveAppResourceId(env, "APP_ID")._unsafeUnwrap()).toBe(resourceId);
    expect(resolveAppResourceId({ TEAMS_APP_ID: "legacy-id" }).isErr()).toBe(flag === "true");
  });

  it("AC-20: flag-off package identity avoids migration YAML lookup", async () => {
    vi.stubEnv(FeatureFlagName.NewDeveloperPortalApis, "false");
    const parse = vi.spyOn(metadataUtil, "parse");
    const yaml = vi.spyOn(pathUtils, "getYmlFilePath");
    expect(
      (
        await resolveProjectAppResourceId(
          "project",
          { APP_ID: resourceId, TEAMS_APP_ID: resourceId },
          manifestId
        )
      )._unsafeUnwrap()
    ).toBe(manifestId);
    expect(yaml).not.toHaveBeenCalled();
    expect(parse).not.toHaveBeenCalled();
  });

  it("AC-20: flag-off safety checks the actual legacy destination for split state", () => {
    vi.stubEnv(FeatureFlagName.NewDeveloperPortalApis, "false");
    const env = { APP_ID: resourceId, APP_MANIFEST_ID: resourceId, TEAMS_APP_ID: manifestId };
    expect(resolveAppResourceId(env).isErr()).toBe(true);
    expect(resolveAppResourceId(env, "APP_ID")._unsafeUnwrap()).toBe(resourceId);
    expect(resolveAppResourceId(env, undefined, manifestId).isErr()).toBe(true);
  });

  it.each([
    { defaultValue: "true", override: undefined, enabled: true },
    { defaultValue: "false", override: undefined, enabled: false },
    { defaultValue: "true", override: "false", enabled: false },
    { defaultValue: "false", override: "true", enabled: true },
  ])(
    "AC-06/AC-07/AC-15: routing, templates, and identity honor the same flag (%j)",
    ({ defaultValue, override, enabled }) => {
      const originalDefault = FeatureFlags.NewDeveloperPortalApis.defaultValue;
      try {
        FeatureFlags.NewDeveloperPortalApis.defaultValue = defaultValue;
        vi.stubEnv(FeatureFlagName.NewDeveloperPortalApis, override);
        expect(Generator.getDefaultVariables("test").NewDeveloperPortalApis).toBe(
          enabled ? "true" : ""
        );
        expect(defaultAppIdEnvName()).toBe(enabled ? "APP_ID" : "TEAMS_APP_ID");
        expect(getActiveTeamsDevPortalClient()).toBe(
          enabled ? devPortalClient : legacyTeamsDevPortalClient
        );
        expect(
          resolveAppResourceId({ APP_ID: resourceId, APP_MANIFEST_ID: manifestId }).isOk()
        ).toBe(enabled);
      } finally {
        FeatureFlags.NewDeveloperPortalApis.defaultValue = originalDefault;
      }
    }
  );

  it.each([undefined, null, "", 42])(
    "AC-10: missing or invalid tenant claim fails explicitly (%j)",
    async (tid) => {
      const provider = new MockedM365Provider();
      vi.spyOn(provider, "getJsonObject").mockResolvedValue(ok({ tid }));
      expect((await getAppTenantId(provider)).isErr()).toBe(true);
    }
  );

  it("AC-10: tenant lookup preserves the authentication error", async () => {
    const provider = new MockedM365Provider();
    const error = new UserError("test", "AuthenticationFailed", "authentication failed");
    vi.spyOn(provider, "getJsonObject").mockResolvedValue(err(error));
    expect((await getAppTenantId(provider))._unsafeUnwrapErr()).toBe(error);
  });

  it("AC-04: project parse errors cannot silently fall back to environment IDs", async () => {
    const error = new UserError("test", "InvalidProject", "invalid project");
    vi.spyOn(pathUtils, "getYmlFilePath").mockReturnValue("m365agents.yml");
    vi.spyOn(metadataUtil, "parse").mockResolvedValue(err(error));
    expect(
      (await resolveProjectAppResourceId("project", { APP_ID: resourceId }))._unsafeUnwrapErr()
    ).toBe(error);
  });
  it("AC-03: falls back to the legacy ID without modifying environment state", () => {
    const env = { APP_ID: "", TEAMS_APP_ID: resourceId };
    expect(resolveAppResourceId(env)._unsafeUnwrap()).toBe(resourceId);
    expect(env).toEqual({ APP_ID: "", TEAMS_APP_ID: resourceId });
  });

  it("AC-04: honors custom mappings then canonical precedence", () => {
    const env = { CUSTOM_ID: manifestId, APP_ID: resourceId, TEAMS_APP_ID: manifestId };
    expect(resolveAppResourceId(env, "CUSTOM_ID")._unsafeUnwrap()).toBe(manifestId);
    expect(resolveAppResourceId(env)._unsafeUnwrap()).toBe(resourceId);
    expect(getAppIdEnvName({ appId: "CUSTOM_ID" })).toBe("CUSTOM_ID");
    expect(getAppIdEnvName({ teamsAppId: "LEGACY_CUSTOM" })).toBe("LEGACY_CUSTOM");
  });

  it("AC-04: does not hide an invalid selected ID with a valid fallback", () => {
    expect(resolveAppResourceId({ APP_ID: "invalid", TEAMS_APP_ID: resourceId }).isErr()).toBe(
      true
    );
  });

  it("AC-15: rejects split-ID state with the new API disabled", () => {
    vi.stubEnv(FeatureFlagName.NewDeveloperPortalApis, "false");
    const result = resolveAppResourceId({ APP_ID: resourceId, APP_MANIFEST_ID: manifestId });
    expect(result.isErr()).toBe(true);
    if (result.isErr()) expect(result.error.message).toContain("TEAMSFX_NEW_DP_APIS");
  });

  it("AC-17: never uses the manifest as a missing split-ID resource", () => {
    vi.stubEnv(FeatureFlagName.NewDeveloperPortalApis, "true");
    expect(
      resolveAppResourceId({ APP_MANIFEST_ID: manifestId }, undefined, manifestId).isErr()
    ).toBe(true);
    expect(resolveAppResourceId({}, undefined, manifestId)._unsafeUnwrap()).toBe(manifestId);
    expect(
      resolveAppResourceId(
        { APP_ID: resourceId, APP_MANIFEST_ID: manifestId },
        undefined,
        manifestId
      )._unsafeUnwrap()
    ).toBe(resourceId);
  });
});

describe("teamsApp/create schema compatibility", () => {
  const validator = new Validator();
  for (const version of validator.supportedVersions()) {
    it(`AC-01: accepts either output key and rejects ambiguity in ${version}`, () => {
      const document = (output: Record<string, string>) => ({
        version,
        provision: [
          { uses: "teamsApp/create", with: { name: "example" }, writeToEnvironmentFile: output },
        ],
      });
      expect(validator.validate(document({ appId: "APP_ID" }), version)).toBe(true);
      expect(validator.validate(document({ teamsAppId: "TEAMS_APP_ID" }), version)).toBe(true);
      expect(
        validator.validate(document({ appId: "APP_ID", teamsAppId: "TEAMS_APP_ID" }), version)
      ).toBe(false);
      expect(validator.validate(document({}), version)).toBe(false);
    });
  }
});
