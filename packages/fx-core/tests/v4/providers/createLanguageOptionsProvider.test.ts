// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { assert, expect } from "vitest";
import { createDefaultCreateOptionsProviders } from "../../../src/v4/providers/createOptionsProviders";
import { resolveLanguageOptions } from "../../../src/v4/providers/createLanguageOptionsProvider";
import { SystemError } from "@microsoft/teamsfx-api";

function languageProvider(
  descriptor: unknown,
  surface = "vscode",
  flagReader: (name: string) => boolean = () => false
) {
  const registry = createDefaultCreateOptionsProviders(
    async () => ({ tools: [], requiresAuth: false }),
    async () => [],
    async () => [],
    { descriptor, surface, flagReader }
  );
  const provider = registry["create.languages"];
  assert.isDefined(provider, "the named language provider must be registered");
  return provider;
}

describe("create language options provider", () => {
  it("CLEAN-02: preserves typed provider errors and wraps unexpected failures", async () => {
    const typed = new SystemError({ source: "test", name: "LanguageLoadError", message: "test" });
    const preserved = await resolveLanguageOptions({
      fetch: () => {
        throw typed;
      },
    });
    assert.strictEqual(preserved._unsafeUnwrapErr(), typed);
    for (const cause of [new Error("provider failure"), "provider failure"]) {
      const result = await resolveLanguageOptions({
        fetch: () => {
          throw cause;
        },
      });
      assert.equal(result._unsafeUnwrapErr().name, "InputProviderFailed");
    }
  });

  it("CLEAN-01: presentation comes from metadata, not template identity", async () => {
    const result = await languageProvider({
      id: "arbitrary-template",
      languages: ["typescript", "python"],
      languageOptions: [{ id: "python", description: "preview.translation.key" }],
    }).fetch({});

    assert.deepEqual(result.options, [
      { id: "typescript", label: "TypeScript" },
      { id: "python", label: "Python", description: "preview.translation.key" },
    ]);
    const legacy = await languageProvider({
      id: "custom-copilot-basic",
      languages: ["python"],
    }).fetch({});
    assert.deepEqual(legacy.options, [{ id: "python", label: "Python" }]);
  });

  it("CLEAN-03: preserves surface gating, order, labels, and unknown ID fallback", async () => {
    const descriptor = { languages: ["javascript", "csharp", "typescript", "rust"] };
    for (const surface of ["vscode", "cli", "vs"]) {
      for (const enabled of [false, true]) {
        const result = await languageProvider(descriptor, surface, (name) => {
          assert.equal(name, "TEAMSFX_CLI_DOTNET");
          return enabled;
        }).fetch({});
        const expected =
          surface !== "vscode" && enabled
            ? ["javascript", "csharp", "typescript", "rust"]
            : ["javascript", "typescript", "rust"];
        assert.deepEqual(
          result.options.map((option) => option.id),
          expected
        );
        assert.equal(result.options.at(-1)?.label, "rust");
      }
    }
  });

  it("CLEAN-07: missing language metadata retains the common fallback", async () => {
    const result = await languageProvider({}).fetch({});
    assert.deepEqual(result.options, [{ id: "common", label: "common" }]);
  });

  for (const overrides of [
    [{ id: "python" }, { id: "python" }],
    [{ id: "csharp" }],
    [{ id: "python", description: false }],
    [{ id: "python", condition: { expr: "true" } }],
  ]) {
    it(`CLEAN-02: rejects invalid language presentation ${JSON.stringify(overrides)}`, async () => {
      await expect(
        async () =>
          await languageProvider({
            languages: ["python"],
            languageOptions: overrides,
          }).fetch({})
      ).rejects.toMatchObject({ name: "TemplateLanguageOptionsInvalid" });
    });
  }
});
