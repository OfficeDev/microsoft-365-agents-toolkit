// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { expect } from "chai";
import { describe, it } from "mocha";
import type { OpenAPIV3 } from "openapi-types";
import {
  createValidator,
  CopilotValidator,
  SMEValidator,
  TeamsAIValidator,
} from "../../../src/specParser/validator";
import { ProjectType, ErrorType } from "../../../src/specParser/types";

// Minimal valid spec for validator tests
function makeSpec(overrides?: Partial<OpenAPIV3.Document>): OpenAPIV3.Document {
  return {
    openapi: "3.0.0",
    info: { title: "Test", version: "1.0.0" },
    servers: [{ url: "https://api.example.com" }],
    paths: {
      "/pets": {
        get: {
          operationId: "listPets",
          summary: "List pets",
          parameters: [
            {
              name: "limit",
              in: "query",
              required: true,
              schema: { type: "integer" },
            },
          ],
          responses: {
            "200": {
              description: "OK",
              content: {
                "application/json": {
                  schema: { type: "array", items: { type: "object" } },
                },
              },
            },
          },
        },
      },
    },
    ...overrides,
  } as OpenAPIV3.Document;
}

describe("specParser/validator", () => {
  describe("createValidator", () => {
    it("creates CopilotValidator for Copilot type", () => {
      const v = createValidator(makeSpec(), { projectType: ProjectType.Copilot });
      expect(v).to.be.instanceOf(CopilotValidator);
    });

    it("creates SMEValidator for SME type", () => {
      const v = createValidator(makeSpec(), { projectType: ProjectType.SME });
      expect(v).to.be.instanceOf(SMEValidator);
    });

    it("creates TeamsAIValidator for TeamsAi type", () => {
      const v = createValidator(makeSpec(), { projectType: ProjectType.TeamsAi });
      expect(v).to.be.instanceOf(TeamsAIValidator);
    });

    it("defaults to SMEValidator", () => {
      const v = createValidator(makeSpec(), {});
      expect(v).to.be.instanceOf(SMEValidator);
    });
  });

  describe("CopilotValidator", () => {
    it("validates a valid spec with no errors", () => {
      const v = createValidator(makeSpec(), {
        projectType: ProjectType.Copilot,
        allowMethods: ["get"],
      });
      const result = v.validateSpec();
      expect(result.errors).to.have.length(0);
    });

    it("validates a valid API", () => {
      const v = createValidator(makeSpec(), {
        projectType: ProjectType.Copilot,
        allowMethods: ["get"],
      });
      const result = v.validateAPI("get", "/pets");
      expect(result.isValid).to.be.true;
    });

    it("rejects missing operationId when allowMissingId is false", () => {
      const spec = makeSpec();
      delete spec.paths["/pets"]!.get!.operationId;
      const v = createValidator(spec, {
        projectType: ProjectType.Copilot,
        allowMethods: ["get"],
        allowMissingId: false,
      });
      const result = v.validateAPI("get", "/pets");
      expect(result.isValid).to.be.false;
      expect(result.reason).to.include(ErrorType.MissingOperationId);
    });
  });

  describe("SMEValidator", () => {
    it("validates spec with version and server checks", () => {
      const v = createValidator(makeSpec(), {
        projectType: ProjectType.SME,
        allowMethods: ["get"],
      });
      const result = v.validateSpec();
      expect(result.errors).to.have.length(0);
    });

    it("detects missing response json", () => {
      const spec = makeSpec();
      spec.paths["/pets"]!.get!.responses = {
        "200": { description: "OK" },
      };
      const v = createValidator(spec, {
        projectType: ProjectType.SME,
        allowMethods: ["get"],
      });
      const result = v.validateAPI("get", "/pets");
      expect(result.isValid).to.be.false;
      expect(result.reason).to.include(ErrorType.ResponseJsonIsEmpty);
    });
  });

  describe("TeamsAIValidator", () => {
    it("validates a valid spec", () => {
      const v = createValidator(makeSpec(), {
        projectType: ProjectType.TeamsAi,
        allowMethods: ["get"],
      });
      const result = v.validateSpec();
      expect(result.errors).to.have.length(0);
    });
  });

  describe("listAPIs", () => {
    it("lists valid APIs from spec", () => {
      const v = createValidator(makeSpec(), {
        projectType: ProjectType.Copilot,
        allowMethods: ["get"],
      });
      const apis = v.listAPIs();
      expect(apis).to.have.property("GET /pets");
      expect(apis["GET /pets"].isValid).to.be.true;
    });

    it("caches API map", () => {
      const v = createValidator(makeSpec(), {
        projectType: ProjectType.Copilot,
        allowMethods: ["get"],
      });
      const first = v.listAPIs();
      const second = v.listAPIs();
      expect(first).to.equal(second); // same reference
    });
  });
});
