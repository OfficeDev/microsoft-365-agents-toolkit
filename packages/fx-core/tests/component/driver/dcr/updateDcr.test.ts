// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { FeatureFlagName } from "../../../../src/common/featureFlags";
import {
  createDcrUpdatePlan,
  validateUpdateDcrArgs,
} from "../../../../src/component/driver/dcr/utility/updateDcr";
import { OauthSupportedAccountTypesUnsupportedCloudError } from "../../../../src/component/driver/oauth/error/oauthSupportedAccountTypesUnsupportedCloud";
import { DcrRegistration } from "../../../../src/component/driver/teamsApp/interfaces/DcrRegistration";
import {
  OauthRegistrationAppType,
  OauthRegistrationSupportedAccountType,
  OauthRegistrationTargetAudience,
} from "../../../../src/component/driver/teamsApp/interfaces/OauthRegistration";

const configurationId = "existing-dcr-configuration";

function currentConfiguration(overrides: Partial<DcrRegistration> = {}): DcrRegistration {
  return {
    clientName: "Existing client",
    m365AppId: "",
    applicableToApps: OauthRegistrationAppType.AnyApp,
    targetAudience: OauthRegistrationTargetAudience.HomeTenant,
    mcpResourceUrl: "https://example.com/mcp",
    wellKnownAuthorizationServer:
      "https://login.example.com/.well-known/oauth-authorization-server",
    resource: "https://resource.example.com/api",
    targetUrlsShouldStartWith: ["https://example.com/mcp"],
    supportedAccountTypes: OauthRegistrationSupportedAccountType.Enterprise,
    ...overrides,
  };
}

function plan(input: Record<string, unknown>, current = currentConfiguration()) {
  return validateUpdateDcrArgs({ configurationId, ...input }).andThen((args) =>
    createDcrUpdatePlan(args, current)
  );
}

describe("DCR update planning (supplementary file-unit coverage)", () => {
  beforeEach(() => {
    vi.stubEnv(FeatureFlagName.SovereignCloudEnvironment, undefined);
  });

  afterEach(() => {
    vi.unstubAllEnvs();
  });

  describe("input validation", () => {
    it("DCR-UPD-03: accepts an ID without inserting defaults or requiring a name", () => {
      const result = validateUpdateDcrArgs({ configurationId });
      expect(result.isOk()).toBe(true);
      if (result.isOk()) {
        expect(result.value).toEqual({ configurationId });
      }
    });

    it("DCR-UPD-04: preserves every valid registration input without inserting fields", () => {
      const input = {
        configurationId,
        name: "n".repeat(128),
        appId: "existing-app",
        mcpResourceUrl: "https://example.com/full/mcp/path",
        wellKnownAuthorizationServer:
          "https://login.example.com/.well-known/oauth-authorization-server",
        resource: "urn:example:resource",
        targetUrlsShouldStartWith: ["https://example.com/full/mcp/path"],
        applicableToApps: OauthRegistrationAppType.SpecificApp,
        targetAudience: OauthRegistrationTargetAudience.AnyTenant,
        includePersonalMicrosoftAccounts: false,
      };
      const result = validateUpdateDcrArgs(input);
      expect(result.isOk()).toBe(true);
      if (result.isOk()) {
        expect(result.value).toEqual(input);
        expect(result.value.targetUrlsShouldStartWith).not.toBe(input.targetUrlsShouldStartWith);
      }
    });

    for (const input of [undefined, null, [], false, 42, "input", {}, { configurationId: "" }]) {
      it(`DCR-UPD-03: rejects invalid input ${JSON.stringify(input)}`, () => {
        const result = validateUpdateDcrArgs(input);
        expect(result.isErr()).toBe(true);
        if (result.isErr()) {
          expect(result.error.name).toBe("InvalidActionInputError");
        }
      });
    }

    const invalidInputs: { title: string; input: Record<string, unknown>; error?: string }[] = [
      { title: "blank ID", input: { configurationId: " " } },
      { title: "non-string ID", input: { configurationId: 42 } },
      { title: "empty name", input: { name: "" } },
      { title: "non-string name", input: { name: 42 } },
      { title: "overlong name", input: { name: "n".repeat(129) }, error: "DcrNameTooLong" },
      { title: "empty app ID", input: { appId: "" } },
      { title: "non-string app ID", input: { appId: 42 } },
      { title: "HTTP MCP endpoint", input: { mcpResourceUrl: "http://example.com/mcp" } },
      { title: "invalid MCP endpoint", input: { mcpResourceUrl: "not-a-url" } },
      {
        title: "invalid metadata URL",
        input: { wellKnownAuthorizationServer: "http://example.com/metadata" },
        error: "DcrWellKnownInvalid",
      },
      { title: "relative resource", input: { resource: "/relative" } },
      { title: "resource fragment", input: { resource: "https://example.com/api#fragment" } },
      { title: "non-string resource", input: { resource: 42 } },
      { title: "unknown app mode", input: { applicableToApps: "AllApps" } },
      { title: "unknown tenant scope", input: { targetAudience: "AllTenants" } },
      { title: "non-array targets", input: { targetUrlsShouldStartWith: "https://example.com" } },
      { title: "non-string target", input: { targetUrlsShouldStartWith: [42] } },
      { title: "HTTP target", input: { targetUrlsShouldStartWith: ["http://example.com"] } },
      { title: "string account option", input: { includePersonalMicrosoftAccounts: "true" } },
      { title: "numeric account option", input: { includePersonalMicrosoftAccounts: 1 } },
      { title: "static client ID", input: { clientId: "static-client" } },
      { title: "static secret", input: { clientSecret: "not-a-real-secret" } },
      { title: "unknown setting", input: { unexpected: true } },
    ];
    for (const field of [
      "name",
      "appId",
      "mcpResourceUrl",
      "wellKnownAuthorizationServer",
      "resource",
      "targetUrlsShouldStartWith",
      "applicableToApps",
      "targetAudience",
      "includePersonalMicrosoftAccounts",
    ]) {
      invalidInputs.push({ title: `null ${field}`, input: { [field]: null } });
    }

    for (const { title, input, error } of invalidInputs) {
      it(`DCR-UPD-04: rejects ${title}`, () => {
        const result = validateUpdateDcrArgs({ configurationId, ...input });
        expect(result.isErr()).toBe(true);
        if (result.isErr()) {
          expect(result.error.name).toBe(error ?? "InvalidActionInputError");
        }
      });
    }

    for (const resource of ["https://example.com/api", "api://resource", "urn:example:resource"]) {
      it(`DCR-UPD-12: accepts an absolute resource URI ${resource}`, () => {
        const result = validateUpdateDcrArgs({ configurationId, resource });
        expect(result.isOk()).toBe(true);
        if (result.isOk()) {
          expect(result.value.resource).toBe(resource);
        }
      });
    }

    for (const includePersonalMicrosoftAccounts of [true, false]) {
      it(`DCR-UPD-18: rejects explicit ${includePersonalMicrosoftAccounts} outside Public cloud`, () => {
        vi.stubEnv(FeatureFlagName.SovereignCloudEnvironment, "GCC H");
        const result = validateUpdateDcrArgs({ configurationId, includePersonalMicrosoftAccounts });
        expect(result.isErr()).toBe(true);
        if (result.isErr()) {
          expect(result.error).toBeInstanceOf(OauthSupportedAccountTypesUnsupportedCloudError);
        }
      });
    }

    it("DCR-UPD-18: omission does not apply the account-option cloud restriction", () => {
      vi.stubEnv(FeatureFlagName.SovereignCloudEnvironment, "GCC H");
      expect(validateUpdateDcrArgs({ configurationId, name: "New name" }).isOk()).toBe(true);
    });
  });

  describe("sparse updates", () => {
    it("DCR-UPD-05: enabling MSA changes only the supported account types", () => {
      const current = currentConfiguration();
      const before = structuredClone(current);
      const result = plan({ includePersonalMicrosoftAccounts: true }, current);
      expect(result.isOk()).toBe(true);
      if (result.isOk()) {
        expect(result.value).toEqual({
          changes: { supportedAccountTypes: "Enterprise, Consumer" },
          clearAppBinding: false,
          discoveryInputsChanged: false,
        });
      }
      expect(current).toEqual(before);
    });

    it("DCR-UPD-06: disabling MSA retains enterprise accounts without changing other settings", () => {
      const result = plan(
        { includePersonalMicrosoftAccounts: false },
        currentConfiguration({
          supportedAccountTypes: OauthRegistrationSupportedAccountType.EnterpriseAndConsumer,
        })
      );
      expect(result.isOk()).toBe(true);
      if (result.isOk()) {
        expect(result.value.changes).toEqual({ supportedAccountTypes: "Enterprise" });
      }
    });

    it("DCR-UPD-07: omission never writes account support, app binding or defaults", () => {
      const result = plan(
        { name: "New name" },
        currentConfiguration({
          applicableToApps: OauthRegistrationAppType.SpecificApp,
          m365AppId: "existing-app",
          supportedAccountTypes: OauthRegistrationSupportedAccountType.EnterpriseAndConsumer,
        })
      );
      expect(result.isOk()).toBe(true);
      if (result.isOk()) {
        expect(result.value.changes).toEqual({ clientName: "New name" });
        expect(result.value.clearAppBinding).toBe(false);
      }
    });

    it("DCR-UPD-09: ID-only input produces an empty change set", () => {
      const result = plan({});
      expect(result.isOk()).toBe(true);
      if (result.isOk()) {
        expect(result.value).toEqual({
          changes: {},
          clearAppBinding: false,
          discoveryInputsChanged: false,
        });
      }
    });

    it("DCR-UPD-09: matching explicit values produce no changes", () => {
      const current = currentConfiguration({
        applicableToApps: OauthRegistrationAppType.SpecificApp,
        m365AppId: "existing-app",
      });
      const result = plan(
        {
          name: current.clientName,
          appId: current.m365AppId,
          mcpResourceUrl: current.mcpResourceUrl,
          wellKnownAuthorizationServer: current.wellKnownAuthorizationServer,
          resource: current.resource,
          targetUrlsShouldStartWith: current.targetUrlsShouldStartWith,
          applicableToApps: current.applicableToApps,
          targetAudience: current.targetAudience,
          includePersonalMicrosoftAccounts: false,
        },
        current
      );
      expect(result.isOk()).toBe(true);
      if (result.isOk()) {
        expect(result.value.changes).toEqual({});
        expect(result.value.clearAppBinding).toBe(false);
        expect(result.value.discoveryInputsChanged).toBe(false);
      }
    });

    it("DCR-UPD-06: an absent readback value is not assumed to be a known enterprise setting", () => {
      const result = plan(
        { includePersonalMicrosoftAccounts: false },
        currentConfiguration({ supportedAccountTypes: undefined })
      );
      expect(result.isOk()).toBe(true);
      if (result.isOk()) {
        expect(result.value.changes).toEqual({ supportedAccountTypes: "Enterprise" });
      }
    });

    it("DCR-UPD-17: tenant-scope changes do not infer other policy changes", () => {
      const result = plan({ targetAudience: OauthRegistrationTargetAudience.AnyTenant });
      expect(result.isOk()).toBe(true);
      if (result.isOk()) {
        expect(result.value.changes).toEqual({ targetAudience: "AnyTenant" });
      }
    });
  });

  describe("discovery and target changes", () => {
    for (const { id, field, value } of [
      { id: "DCR-UPD-10", field: "mcpResourceUrl", value: "https://new.example.com/full/mcp/path" },
      {
        id: "DCR-UPD-11",
        field: "wellKnownAuthorizationServer",
        value: "https://new.example.com/.well-known/oauth-authorization-server",
      },
      { id: "DCR-UPD-12", field: "resource", value: "urn:example:new-resource" },
    ]) {
      it(`${id}: plans only the explicit ${field} change without inventing discovered values`, () => {
        const result = plan({ [field]: value });
        expect(result.isOk()).toBe(true);
        if (result.isOk()) {
          expect(result.value.changes).toEqual({ [field]: value });
          expect(result.value.discoveryInputsChanged).toBe(true);
        }
      });
    }

    it("DCR-UPD-12: discovery changes require an effective discovery source", () => {
      const result = plan(
        { resource: "urn:example:new-resource" },
        currentConfiguration({ mcpResourceUrl: undefined, wellKnownAuthorizationServer: undefined })
      );
      expect(result.isErr()).toBe(true);
      if (result.isErr()) {
        expect(result.error.name).toBe("InvalidActionInputError");
      }
    });

    it("DCR-UPD-07: a policy change does not revalidate unrelated discovery inputs", () => {
      const result = plan(
        { includePersonalMicrosoftAccounts: true },
        currentConfiguration({ mcpResourceUrl: undefined, wellKnownAuthorizationServer: undefined })
      );
      expect(result.isOk()).toBe(true);
    });

    for (const targetUrlsShouldStartWith of [
      [],
      ["https://new.example.com/mcp"],
      ["https://one.example.com/mcp", "https://two.example.com/mcp"],
    ]) {
      it(`DCR-UPD-13: uses replacement semantics for ${targetUrlsShouldStartWith.length} target entries`, () => {
        const result = plan({ targetUrlsShouldStartWith });
        expect(result.isOk()).toBe(true);
        if (result.isOk()) {
          expect(result.value.changes).toEqual({ targetUrlsShouldStartWith });
          expect(result.value.discoveryInputsChanged).toBe(false);
        }
      });
    }

    it("DCR-UPD-09: target ordering alone does not create a change or mutate arrays", () => {
      const targets = ["https://one.example.com/mcp", "https://two.example.com/mcp"];
      const current = currentConfiguration({ targetUrlsShouldStartWith: [...targets].reverse() });
      const before = structuredClone(current);
      const result = plan({ targetUrlsShouldStartWith: targets }, current);
      expect(result.isOk()).toBe(true);
      if (result.isOk()) {
        expect(result.value.changes).toEqual({});
      }
      expect(current).toEqual(before);
      expect(targets[0]).toBe("https://one.example.com/mcp");
    });

    it("DCR-UPD-13: the planned target array does not alias the caller's input", () => {
      const targets = ["https://new.example.com/mcp"];
      const result = plan({ targetUrlsShouldStartWith: targets });
      targets.push("https://later.example.com/mcp");
      expect(result.isOk()).toBe(true);
      if (result.isOk()) {
        expect(result.value.changes.targetUrlsShouldStartWith).toEqual([
          "https://new.example.com/mcp",
        ]);
      }
    });
  });

  describe("application restrictions", () => {
    it("DCR-UPD-14: changes an existing SpecificApp binding without resupplying the mode", () => {
      const result = plan(
        { appId: "new-app" },
        currentConfiguration({
          applicableToApps: OauthRegistrationAppType.SpecificApp,
          m365AppId: "old-app",
        })
      );
      expect(result.isOk()).toBe(true);
      if (result.isOk()) {
        expect(result.value.changes).toEqual({ m365AppId: "new-app" });
        expect(result.value.clearAppBinding).toBe(false);
      }
    });

    it("DCR-UPD-15: switches to SpecificApp with a supplied binding", () => {
      const result = plan({
        applicableToApps: OauthRegistrationAppType.SpecificApp,
        appId: "new-app",
      });
      expect(result.isOk()).toBe(true);
      if (result.isOk()) {
        expect(result.value.changes).toEqual({
          applicableToApps: "SpecificApp",
          m365AppId: "new-app",
        });
      }
    });

    it("DCR-UPD-15: retains an existing binding when switching to SpecificApp", () => {
      const result = plan(
        { applicableToApps: OauthRegistrationAppType.SpecificApp },
        currentConfiguration({ m365AppId: "existing-app" })
      );
      expect(result.isOk()).toBe(true);
      if (result.isOk()) {
        expect(result.value.changes).toEqual({ applicableToApps: "SpecificApp" });
      }
    });

    it("DCR-UPD-15: rejects an effective SpecificApp configuration without an app ID", () => {
      const result = plan({ applicableToApps: OauthRegistrationAppType.SpecificApp });
      expect(result.isErr()).toBe(true);
      if (result.isErr()) {
        expect(result.error.name).toBe("InvalidActionInputError");
      }
    });

    it("DCR-UPD-16: represents binding removal as intent rather than inventing a wire clear value", () => {
      const result = plan(
        { applicableToApps: OauthRegistrationAppType.AnyApp },
        currentConfiguration({
          applicableToApps: OauthRegistrationAppType.SpecificApp,
          m365AppId: "existing-app",
        })
      );
      expect(result.isOk()).toBe(true);
      if (result.isOk()) {
        expect(result.value.changes).toEqual({ applicableToApps: "AnyApp" });
        expect(result.value.clearAppBinding).toBe(true);
        expect(result.value.changes).not.toHaveProperty("m365AppId");
      }
    });

    it("DCR-UPD-16: appId has no effect while the effective mode is AnyApp", () => {
      const result = plan({ appId: "ignored-app" });
      expect(result.isOk()).toBe(true);
      if (result.isOk()) {
        expect(result.value.changes).toEqual({});
        expect(result.value.clearAppBinding).toBe(false);
      }
    });
  });
});
