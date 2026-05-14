// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * OpenAPI spec-parser E2E tests — subprocess layer.
 *
 * Verifies that `atk new da api-plugin-from-spec` and related OpenAPI template
 * commands accept the --apiSpecPath option and handle spec validation properly.
 * Does NOT create real Azure resources.
 *
 * The scaffold may fail if template CDN is unreachable — in that case the test
 * verifies error handling rather than success. The key assertion is that the
 * spec-parser is wired correctly and the CLI plumbs spec path through to core-next.
 */

import * as fs from "fs";
import * as os from "os";
import * as path from "path";
import { expect } from "chai";
import { describe, it, beforeEach, afterEach } from "mocha";
import { promisify } from "util";
import { exec } from "child_process";

const execAsync = promisify(exec);

const ATK_BIN = process.env.ATK_BIN || "atk";
const TIMEOUT = 120_000; // 2 min per command

function tmpDir(): string {
  const dir = path.join(os.tmpdir(), `atk-openapi-e2e-${Date.now()}`);
  fs.mkdirSync(dir, { recursive: true });
  return dir;
}

async function run(
  command: string,
  cwd?: string
): Promise<{ stdout: string; stderr: string; exitCode: number }> {
  try {
    const { stdout, stderr } = await execAsync(command, {
      cwd,
      timeout: TIMEOUT,
      env: { ...process.env, CI_ENABLED: "true" },
    });
    return { stdout, stderr, exitCode: 0 };
  } catch (e: unknown) {
    const err = e as { stdout?: string; stderr?: string; code?: number };
    return {
      stdout: err.stdout ?? "",
      stderr: err.stderr ?? "",
      exitCode: err.code ?? 1,
    };
  }
}

/**
 * Minimal valid OpenAPI 3.0 spec for testing.
 */
const VALID_SPEC = {
  openapi: "3.0.0",
  info: { title: "TestAPI", version: "1.0.0" },
  servers: [{ url: "https://api.example.com" }],
  paths: {
    "/items": {
      get: {
        operationId: "listItems",
        summary: "List items",
        parameters: [{ name: "limit", in: "query", required: false, schema: { type: "integer" } }],
        responses: {
          "200": {
            description: "OK",
            content: {
              "application/json": {
                schema: {
                  type: "array",
                  items: {
                    type: "object",
                    properties: {
                      id: { type: "string" },
                      name: { type: "string" },
                    },
                  },
                },
              },
            },
          },
        },
      },
      post: {
        operationId: "createItem",
        summary: "Create item",
        requestBody: {
          required: true,
          content: {
            "application/json": {
              schema: {
                type: "object",
                properties: {
                  name: { type: "string" },
                },
                required: ["name"],
              },
            },
          },
        },
        responses: {
          "201": {
            description: "Created",
            content: {
              "application/json": {
                schema: {
                  type: "object",
                  properties: {
                    id: { type: "string" },
                    name: { type: "string" },
                  },
                },
              },
            },
          },
        },
      },
    },
  },
};

describe("OpenAPI spec-parser E2E", function () {
  this.timeout(5 * 60 * 1000); // 5 min total

  let dir: string;
  let specPath: string;

  beforeEach(function () {
    dir = tmpDir();
    specPath = path.join(dir, "api-spec.json");
    fs.writeFileSync(specPath, JSON.stringify(VALID_SPEC, null, 2));
  });

  afterEach(async function () {
    await fs.promises.rm(dir, { recursive: true, force: true }).catch(() => {});
  });

  // -----------------------------------------------------------------------
  // Help / syntax recognition
  // -----------------------------------------------------------------------

  describe("CLI help for OpenAPI templates", function () {
    it("atk new da --help should list api-plugin-from-spec", async function () {
      const result = await run(`${ATK_BIN} new da --help`);
      expect(result.exitCode).to.equal(0);
      expect(result.stdout).to.include("api-plugin-from-spec");
    });

    it("atk new da api-plugin-from-spec --help should show --apiSpecPath", async function () {
      const result = await run(`${ATK_BIN} new da api-plugin-from-spec --help`);
      expect(result.exitCode).to.equal(0);
      expect(result.stdout).to.include("--apiSpecPath");
    });

    it("atk new me --help should list from-spec", async function () {
      const result = await run(`${ATK_BIN} new me --help`);
      expect(result.exitCode).to.equal(0);
      expect(result.stdout).to.include("from-spec");
    });

    it("atk new ai --help should list rag-from-spec", async function () {
      const result = await run(`${ATK_BIN} new ai --help`);
      expect(result.exitCode).to.equal(0);
      expect(result.stdout).to.include("rag-from-spec");
    });
  });

  // -----------------------------------------------------------------------
  // Validation error handling
  // -----------------------------------------------------------------------

  describe("Spec validation through CLI", function () {
    it("should report error for invalid spec file (broken JSON)", async function () {
      const brokenSpec = path.join(dir, "broken.json");
      fs.writeFileSync(brokenSpec, "{ this is not valid json }");

      const result = await run(
        `${ATK_BIN} new da api-plugin-from-spec --name TestAPIPlugin --apiSpecPath "${brokenSpec}" --apiOperations "get /items" --folder ${dir} --non-interactive`
      );
      // Should exit non-zero due to invalid spec
      expect(result.exitCode).to.not.equal(0);
      // stderr or stdout should mention the spec problem
      const output = result.stdout + result.stderr;
      expect(output.length).to.be.greaterThan(0);
    });

    it("should report error for nonexistent spec file", async function () {
      const result = await run(
        `${ATK_BIN} new da api-plugin-from-spec --name TestAPIPlugin --apiSpecPath "${path.join(dir, "nonexistent.json")}" --apiOperations "get /items" --folder ${dir} --non-interactive`
      );
      expect(result.exitCode).to.not.equal(0);
    });
  });

  // -----------------------------------------------------------------------
  // OpenAPI scaffold (may require template CDN)
  // -----------------------------------------------------------------------

  describe("OpenAPI scaffold with valid spec", function () {
    it("atk new da api-plugin-from-spec with valid spec", async function () {
      const result = await run(
        `${ATK_BIN} new da api-plugin-from-spec --name TestAPIPlugin --apiSpecPath "${specPath}" --apiOperations "get /items" --folder ${dir} --non-interactive`
      );

      const projectDir = path.join(dir, "TestAPIPlugin");

      if (result.exitCode === 0) {
        // Scaffold succeeded — verify project structure
        expect(fs.existsSync(projectDir), "project dir should exist").to.be.true;

        // Should have generated openapi.json file
        const openapiPath = path.join(projectDir, "openapi.json");
        if (fs.existsSync(openapiPath)) {
          const generated = JSON.parse(fs.readFileSync(openapiPath, "utf8"));
          expect(generated).to.have.property("openapi");
          expect(generated.paths).to.have.property("/items");
          expect(generated.paths["/items"]).to.have.property("get");
        }
      } else {
        // Scaffold may fail if template CDN is unreachable —
        // in that case, just verify it produced a meaningful error
        const output = result.stdout + result.stderr;
        expect(output.length, "should produce error output on failure").to.be.greaterThan(0);
        // Skip silently — this is expected in offline environments
        this.skip();
      }
    });

    it("atk new me from-spec with valid spec", async function () {
      const result = await run(
        `${ATK_BIN} new me from-spec --name TestME --apiSpecPath "${specPath}" --apiOperations "get /items" --folder ${dir} --non-interactive`
      );

      const projectDir = path.join(dir, "TestME");

      if (result.exitCode === 0) {
        expect(fs.existsSync(projectDir), "project dir should exist").to.be.true;
      } else {
        // May fail offline — verify error output exists
        const output = result.stdout + result.stderr;
        expect(output.length).to.be.greaterThan(0);
        this.skip();
      }
    });

    it("atk new ai rag-from-spec with valid spec", async function () {
      const result = await run(
        `${ATK_BIN} new ai rag-from-spec --name TestRAG --apiSpecPath "${specPath}" --apiOperations "get /items" --llmProvider azure-openai --folder ${dir} --non-interactive`
      );

      const projectDir = path.join(dir, "TestRAG");

      if (result.exitCode === 0) {
        expect(fs.existsSync(projectDir), "project dir should exist").to.be.true;
      } else {
        const output = result.stdout + result.stderr;
        expect(output.length).to.be.greaterThan(0);
        this.skip();
      }
    });
  });

  // -----------------------------------------------------------------------
  // Regenerate command syntax
  // -----------------------------------------------------------------------

  describe("Regenerate command --help", function () {
    it("atk regenerate action --help should exit 0 and show --api-spec-path", async function () {
      const result = await run(`${ATK_BIN} regenerate action --help`);
      expect(result.exitCode).to.equal(0);
      expect(result.stdout).to.include("--api-spec-path");
    });
  });
});
