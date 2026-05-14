/**
 * Copyright (c) Microsoft Corporation.
 * Licensed under the MIT license.
 *
 * Integration tests for feature gaps 2-5: publishAppPackage driver,
 * M365PackageService, foundry template, and MetaOS upgrade template.
 *
 * Tests run real core-next registrations in-process with the CLI
 * Commander tree. Only telemetry and console I/O are stubbed.
 */

import { expect } from "chai";
import { describe, it, afterEach, beforeEach } from "mocha";
import * as sinon from "sinon";
import { buildProgram } from "../../src/commands";
import { cliTelemetry } from "../../src/telemetry";
import {
  registerBuiltinDrivers,
  builtinDrivers,
  driverRegistry,
  registerBuiltinTemplates,
  templateRegistry,
  clients,
  publishOp,
} from "@microsoft/teamsfx-core-next";

describe("Gap Features Integration", () => {
  const sandbox = sinon.createSandbox();

  beforeEach(() => {
    sandbox.stub(cliTelemetry, "sendEvent");
    sandbox.stub(cliTelemetry, "sendErrorEvent");
    sandbox.stub(cliTelemetry, "flush").resolves();
    sandbox.stub(console, "log");
    sandbox.stub(console, "warn");
    sandbox.stub(console, "error");
  });

  afterEach(() => {
    sandbox.restore();
    process.exitCode = undefined;
  });

  // ---------------------------------------------------------------------------
  // Gap 2: publishAppPackage driver → Graph API
  // ---------------------------------------------------------------------------
  describe("Gap 2: publishAppPackage driver registration", () => {
    before(() => {
      registerBuiltinDrivers();
    });

    it("publishAppPackage driver should be in builtinDrivers", () => {
      const ids = builtinDrivers.map((d) => d.id);
      expect(ids).to.include("teamsApp/publishAppPackage");
    });

    it("publishAppPackage driver should be registered in driverRegistry", () => {
      expect(driverRegistry.has("teamsApp/publishAppPackage")).to.be.true;
    });

    it("publishAppPackage driver should have execute method", () => {
      const driver = driverRegistry.get("teamsApp/publishAppPackage");
      expect(driver).to.exist;
      expect(driver!.executeFn).to.be.a("function");
    });
  });

  // ---------------------------------------------------------------------------
  // Gap 3: M365PackageService + GraphApiClient exports
  // ---------------------------------------------------------------------------
  describe("Gap 3: M365PackageService and GraphApiClient exports", () => {
    it("GraphApiClient should be exported from clients namespace", () => {
      expect(clients.GraphApiClient).to.be.a("function");
    });

    it("M365PackageService should be exported from clients namespace", () => {
      expect(clients.M365PackageService).to.be.a("function");
    });

    it("mosServiceScopes should be exported", () => {
      expect(clients.mosServiceScopes).to.exist;
      expect(clients.mosServiceScopes).to.be.a("function");
    });
  });

  // ---------------------------------------------------------------------------
  // Gap 4: foundry-agent-to-m365 template → CLI command tree
  // ---------------------------------------------------------------------------
  describe("Gap 4: Foundry template CLI command", () => {
    it("atk new ai should have foundry-to-m365 subcommand", () => {
      const program = buildProgram("atk");
      program.exitOverride();

      const newCmd = program.commands.find((c) => c.name() === "new");
      expect(newCmd, "new command should exist").to.exist;

      const aiCmd = newCmd!.commands.find((c) => c.name() === "ai");
      expect(aiCmd, "ai subcommand should exist").to.exist;

      const templateNames = aiCmd!.commands.map((c) => c.name());
      expect(templateNames).to.include("foundry-to-m365");
    });

    it("foundry-to-m365 should have --foundryEndpoint and --foundryAgentId options", () => {
      const program = buildProgram("atk");
      program.exitOverride();

      const newCmd = program.commands.find((c) => c.name() === "new");
      const aiCmd = newCmd!.commands.find((c) => c.name() === "ai");
      const foundryCmd = aiCmd!.commands.find((c) => c.name() === "foundry-to-m365");
      expect(foundryCmd, "foundry-to-m365 command should exist").to.exist;

      const optNames = foundryCmd!.options.map((o) => o.long);
      expect(optNames).to.include("--foundryEndpoint");
      expect(optNames).to.include("--foundryAgentId");
    });

    it("foundry-to-m365 should have --name (required) and --folder options", () => {
      const program = buildProgram("atk");
      program.exitOverride();

      const newCmd = program.commands.find((c) => c.name() === "new");
      const aiCmd = newCmd!.commands.find((c) => c.name() === "ai");
      const foundryCmd = aiCmd!.commands.find((c) => c.name() === "foundry-to-m365");
      expect(foundryCmd).to.exist;

      const optNames = foundryCmd!.options.map((o) => o.long);
      expect(optNames).to.include("--name");
      expect(optNames).to.include("--folder");
    });

    it("foundry template should be registered in templateRegistry", () => {
      registerBuiltinTemplates();
      const foundry = templateRegistry.listByCategory("ai-agent");
      const ids = foundry.map((d) => d.id);
      expect(ids).to.include("ai-agent/foundry-to-m365");
    });
  });

  // ---------------------------------------------------------------------------
  // Gap 5: da-meta-os-upgrade template → CLI command tree
  // ---------------------------------------------------------------------------
  describe("Gap 5: MetaOS upgrade template CLI command", () => {
    it("atk new da should have metaos-upgrade subcommand", () => {
      const program = buildProgram("atk");
      program.exitOverride();

      const newCmd = program.commands.find((c) => c.name() === "new");
      const daCmd = newCmd!.commands.find((c) => c.name() === "da");
      expect(daCmd, "da subcommand should exist").to.exist;

      const templateNames = daCmd!.commands.map((c) => c.name());
      expect(templateNames).to.include("metaos-upgrade");
    });

    it("metaos-upgrade should have --officeAddinFolder option", () => {
      const program = buildProgram("atk");
      program.exitOverride();

      const newCmd = program.commands.find((c) => c.name() === "new");
      const daCmd = newCmd!.commands.find((c) => c.name() === "da");
      const upgradeCmd = daCmd!.commands.find((c) => c.name() === "metaos-upgrade");
      expect(upgradeCmd, "metaos-upgrade command should exist").to.exist;

      const optNames = upgradeCmd!.options.map((o) => o.long);
      expect(optNames).to.include("--officeAddinFolder");
    });

    it("metaos-upgrade should have --name (required) and --folder options", () => {
      const program = buildProgram("atk");
      program.exitOverride();

      const newCmd = program.commands.find((c) => c.name() === "new");
      const daCmd = newCmd!.commands.find((c) => c.name() === "da");
      const upgradeCmd = daCmd!.commands.find((c) => c.name() === "metaos-upgrade");
      expect(upgradeCmd).to.exist;

      const optNames = upgradeCmd!.options.map((o) => o.long);
      expect(optNames).to.include("--name");
      expect(optNames).to.include("--folder");
    });

    it("metaos-upgrade should be registered in DA category behind DAMetaOS flag", () => {
      registerBuiltinTemplates();
      const daTemplates = templateRegistry.listByCategory("declarative-agent");
      const upgrade = daTemplates.find((d) => d.id === "da/metaos-upgrade");
      expect(upgrade, "da/metaos-upgrade should be registered").to.exist;
      expect(upgrade!.featureFlag).to.equal("DAMetaOS");
    });
  });

  // ---------------------------------------------------------------------------
  // Cross-gap: publish command exists in CLI tree
  // ---------------------------------------------------------------------------
  describe("Publish command integration", () => {
    it("atk publish command should exist", () => {
      const program = buildProgram("atk");
      program.exitOverride();

      const publishCmd = program.commands.find((c) => c.name() === "publish");
      expect(publishCmd, "publish command should exist").to.exist;
    });

    it("atk publish should have --env and --project-folder options", () => {
      const program = buildProgram("atk");
      program.exitOverride();

      const publishCmd = program.commands.find((c) => c.name() === "publish");
      expect(publishCmd).to.exist;

      const optNames = publishCmd!.options.map((o) => o.long);
      expect(optNames).to.include("--env");
      expect(optNames).to.include("--project-folder");
    });
  });

  // ---------------------------------------------------------------------------
  // Gap 2 (extended): publishOp and unpublishTeamsApp
  // ---------------------------------------------------------------------------
  describe("Gap 2: publishOp and unpublish cleanup", () => {
    it("publishOp should be a defined operation", () => {
      expect(publishOp).to.exist;
      expect(publishOp).to.have.property("name", "publish");
    });

    it("GraphApiClient should have unpublishTeamsApp method", () => {
      expect(clients.GraphApiClient.prototype.unpublishTeamsApp).to.be.a("function");
    });

    it("graphAppCatalogScopes should include AppCatalog.ReadWrite.All", () => {
      const scopes = clients.graphAppCatalogScopes();
      expect(scopes).to.include("AppCatalog.ReadWrite.All");
    });
  });
});
