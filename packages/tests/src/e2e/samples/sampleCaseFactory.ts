// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Ivan Chen <v-ivanchen@microsoft.com>
 */

import { expect } from "chai";
import * as fs from "fs-extra";
import * as path from "path";
import { it } from "@microsoft/extra-shot-mocha";
import {
  getTestFolder,
  getUniqueAppName,
  readContextMultiEnvV3,
  validateTabAndBotProjectProvision,
  createResourceGroup,
} from "../commonUtils";
import { Executor } from "../../utils/executor";
import { Cleaner } from "../../commonlib/cleaner";
import { TemplateProjectFolder } from "../../utils/constants";
import { environmentNameManager } from "@microsoft/teamsfx-core";
import {
  AadValidator,
  BotValidator,
  FunctionValidator,
  ContainerAppValidator,
  StaticSiteValidator,
} from "../../commonlib";
import m365Login from "@microsoft/m365agentstoolkit-cli/src/commonlib/m365Login";

export abstract class CaseFactory {
  public sampleName: TemplateProjectFolder;
  public testPlanCaseId: number;
  public author: string;
  public validate: (
    | "bot"
    | "tab"
    | "aad"
    | "dashboard"
    | "sql"
    | "function"
    | "spfx"
    | "tab & bot"
    | "aca"
  )[] = [];
  public options?: {
    skipProvision?: boolean;
    skipDeploy?: boolean;
    skipValidate?: boolean;
    skipPackage?: boolean;
    skipPreview?: boolean;
    manifestFolderName?: string;
  };

  public constructor(
    sampleName: TemplateProjectFolder,
    testPlanCaseId: number,
    author: string,
    validate: (
      | "bot"
      | "tab"
      | "aad"
      | "dashboard"
      | "sql"
      | "function"
      | "spfx"
      | "tab & bot"
      | "aca"
    )[] = [],
    options: {
      skipProvision?: boolean;
      skipDeploy?: boolean;
      skipValidate?: boolean;
      skipPackage?: boolean;
      manifestFolderName?: string;
    } = {}
  ) {
    this.sampleName = sampleName;
    this.testPlanCaseId = testPlanCaseId;
    this.author = author;
    this.validate = validate;
    this.options = options;
  }

  public onBefore(): Promise<void> {
    return Promise.resolve();
  }

  public async onAfter(projectPath: string): Promise<void> {
    await Cleaner.clean(projectPath);
  }

  public async onAfterCreate(projectPath: string): Promise<void> {
    expect(fs.pathExistsSync(path.resolve(projectPath, "infra"))).to.be.true;
  }

  public async onCreate(
    appName: string,
    testFolder: string,
    sampleName: TemplateProjectFolder
  ): Promise<void> {
    await Executor.openTemplateProject(appName, testFolder, sampleName);
  }

  public async onBeforeProvision(projectPath: string): Promise<void> {
    return Promise.resolve();
  }

  public test() {
    const {
      sampleName,
      testPlanCaseId,
      author,
      validate,
      options,
      onBefore,
      onAfter,
      onAfterCreate,
      onBeforeProvision,
      onCreate,
    } = this;
    describe(`Sample Tests: ${sampleName}`, function () {
      const testFolder = getTestFolder();
      const appName = getUniqueAppName();
      const projectPath = path.resolve(testFolder, appName);
      const env = environmentNameManager.getDefaultEnvName();
      before(async () => {
        await onBefore();
      });

      after(async function () {
        await onAfter(projectPath);
      });

      it(sampleName, { testPlanCaseId, author }, async function () {
        // create project
        await onCreate(appName, testFolder, sampleName);
        expect(fs.pathExistsSync(projectPath)).to.be.true;

        await onAfterCreate(projectPath);

        // provision
        {
          if (options?.skipProvision) {
            console.log("skip Provision...");
            console.log("debug finish!");
            return;
          }

          await onBeforeProvision(projectPath);

          const result = await createResourceGroup(appName + "-rg", "westus");
          expect(result).to.be.true;
          process.env["AZURE_RESOURCE_GROUP_NAME"] = appName + "-rg";

          const { success } = await Executor.provision(projectPath);
          expect(success).to.be.true;

          // Validate Provision
          const context = await readContextMultiEnvV3(projectPath, env);
          if (validate.includes("bot")) {
            // Validate Bot Provision
            const bot = new BotValidator(context, projectPath, env);
            await bot.validateProvisionV3(false);
          }
          if (validate.includes("aad")) {
            // Validate Aad App
            const aad = AadValidator.init(context, false, m365Login);
            await AadValidator.validate(aad);
          }
          if (validate.includes("tab & bot")) {
            // Validate Tab & Bot Provision
            await validateTabAndBotProjectProvision(projectPath, env);
          }
          if (validate.includes("function")) {
            // Validate Function App
            const functionValidator = new FunctionValidator(
              context,
              projectPath,
              env
            );
            await functionValidator.validateProvision();
          }
          if (validate.includes("aca")) {
            // Validate Container App Provision
            const aca = new ContainerAppValidator(context);
            await aca.validateProvision(false);
          }
        }

        // deploy
        {
          if (options?.skipDeploy) {
            console.log("skip Deploy...");
            console.log("debug finish!");
            return;
          }

          if (validate.includes("aca")) {
            const { success } = await Executor.login();
            expect(success).to.be.true;
          }

          const { success } = await Executor.deploy(projectPath);
          expect(success).to.be.true;

          // Validate deployment
          const context = await readContextMultiEnvV3(projectPath, env);
          if (validate.includes("bot")) {
            // Validate Bot Deploy
            const bot = new BotValidator(context, projectPath, env);
            await bot.validateDeploy();
          }
          if (validate.includes("aca")) {
            const aca = new ContainerAppValidator(context);
            await aca.validateContainerAppStatus();
          }
          if (validate.includes("tab")) {
            const staticSite = StaticSiteValidator.init(context);
            await StaticSiteValidator.validateDeploy(staticSite);
          }
        }

        // validate
        {
          if (options?.skipValidate) {
            console.log("skip Validate...");
            console.log("debug finish!");
            return;
          }
          const { success } = await Executor.validate(
            projectPath,
            "dev",
            options?.manifestFolderName
          );
          expect(success).to.be.true;
        }

        // package
        {
          if (options?.skipPackage) {
            console.log("skip Package...");
            console.log("debug finish!");
            return;
          }
          const { success } = await Executor.package(
            projectPath,
            "dev",
            options?.manifestFolderName
          );
          expect(success).to.be.true;
        }

        // preview
        {
          if (options?.skipPreview) {
            console.log("skip Preview...");
            console.log("debug finish!");
            return;
          }
          const { success } = await Executor.preview(projectPath);
          expect(success).to.be.true;
        }
      });
    });
  }
}
