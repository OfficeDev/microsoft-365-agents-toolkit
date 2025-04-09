// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import "mocha";
import * as chai from "chai";
import * as sinon from "sinon";
import {
  migrateTransparentLocalTunnel,
  migratePrepareManifest,
  migrateSetUpBot,
  migrateSetUpSSO,
  migrateSetUpTab,
  migrateTransparentNpmInstall,
  migrateTransparentPrerequisite,
  migrateNgrokStartTask,
  migrateNgrokStartCommand,
  migrateAuthStart,
  migrateBackendWatch,
  migrateFrontendStart,
  migrateBackendStart,
  migrateGetFuncPathCommand,
} from "../../../../src/core/middleware/utils/debug/taskMigrator";
import { CommentArray, CommentJSONValue, parse, stringify } from "comment-json";
import { DebugMigrationContext } from "../../../../src/core/middleware/utils/debug/debugMigrationContext";
import * as debugV3MigrationUtils from "../../../../src/core/middleware/utils/debug/debugV3MigrationUtils";
import { ok } from "@microsoft/teamsfx-api";
import { LocalCrypto } from "../../../../src/core/crypto";
import { mockMigrationContext } from "./utils";
import * as os from "os";
import * as path from "path";
import { NodeChecker } from "../../../../src/component/deps-checker/internal/nodeChecker";

describe("debugMigration", () => {
  const projectPath = ".";

  describe("migrateTransparentPrerequisite", () => {
    afterEach(async () => {
      sinon.restore();
    });
    it("happy path", async () => {
      sinon
        .stub(NodeChecker, "getInstalledNodeVersion")
        .resolves({ version: "14.0.0", majorVersion: "14" });
      const migrationContext = await mockMigrationContext(projectPath);
      const testTaskContent = `[
      {
        // Check if all required prerequisites are installed and will install them if not.
        // See https://aka.ms/teamsfx-check-prerequisites-task to know the details and how to customize the args.
        "label": "Validate & install prerequisites",
        "type": "teamsfx",
        "command": "debug-check-prerequisites",
        "args": {
            "prerequisites": [
                "nodejs", // Validate if Node.js is installed.
                "m365Account", // Sign-in prompt for Microsoft 365 account, then validate if the account enables the sideloading permission.
                "devCert", // Install localhost SSL certificate. It's used to serve the development sites over HTTPS to debug the Tab app in Teams.
                "func", // Install Azure Functions Core Tools. It's used to serve Azure Functions hosted project locally.
                "dotnet", // Ensure .NET Core SDK is installed. TeamsFx Azure Functions project depends on extra .NET binding extensions for HTTP trigger authorization.
                "ngrok", // Install Ngrok. Bot project requires a public message endpoint, and ngrok can help create public tunnel for your local service.
                "portOccupancy" // Validate available ports to ensure those debug ones are not occupied.
            ],
            "portOccupancy": [
                53000, // tab service port
                7071, // backend service port
                9229, // backend inspector port for Node.js debugger
                3978, // bot service port
                9239 // bot inspector port for Node.js debugger
            ]
        }
      }]`;
      const expectedTaskContent = `[
        {
          // Check if all required prerequisites are installed and will install them if not.
          // See https://aka.ms/teamsfx-check-prerequisites-task to know the details and how to customize the args.
          "label": "Validate & install prerequisites",
          "type": "teamsfx",
          "command": "debug-check-prerequisites",
          "args": {
              "prerequisites": [
                  "nodejs", // Validate if Node.js is installed.
                  "m365Account", // Sign-in prompt for Microsoft 365 account, then validate if the account enables the sideloading permission.
                  "portOccupancy" // Validate available ports to ensure those debug ones are not occupied.
              ],
              "portOccupancy": [
                  53000, // tab service port
                  7071, // backend service port
                  9229, // backend inspector port for Node.js debugger
                  3978, // bot service port
                  9239 // bot inspector port for Node.js debugger
              ]
          }
        }]`;
      const testTasks = parse(testTaskContent) as CommentArray<CommentJSONValue>;
      const oldProjectSettings = {} as any;
      const debugContext = new DebugMigrationContext(
        migrationContext,
        testTasks,
        oldProjectSettings,
        {}
      );
      await migrateTransparentPrerequisite(debugContext);
      chai.assert.equal(
        stringify(testTasks, null, 4),
        stringify(parse(expectedTaskContent), null, 4)
      );
      chai.assert.isTrue(debugContext.appYmlConfig.deploy?.tools?.devCert?.trust);
      chai.assert.isTrue(debugContext.appYmlConfig.deploy?.tools?.dotnet);
      chai.assert.deepEqual(debugContext.appYmlConfig.deploy?.tools?.func, {
        version: "4",
        symlinkDir: "./devTools/func",
      });
    });

    it("customized prerequisite", async () => {
      const migrationContext = await mockMigrationContext(projectPath);
      const testTaskContent = `[
      {
        // Check if all required prerequisites are installed and will install them if not.
        // See https://aka.ms/teamsfx-check-prerequisites-task to know the details and how to customize the args.
        "label": "Validate & install prerequisites",
        "type": "teamsfx",
        "command": "debug-check-prerequisites",
        "args": {
            "prerequisites": [
                "nodejs", // Validate if Node.js is installed.
                "dotnet", // Ensure .NET Core SDK is installed. TeamsFx Azure Functions project depends on extra .NET binding extensions for HTTP trigger authorization.
                "ngrok", // Install Ngrok. Bot project requires a public message endpoint, and ngrok can help create public tunnel for your local service.
                "portOccupancy" // Validate available ports to ensure those debug ones are not occupied.
            ],
            "portOccupancy": [
                53000, // tab service port
                7071, // backend service port
                9229, // backend inspector port for Node.js debugger
                3978, // bot service port
                9239 // bot inspector port for Node.js debugger
            ]
        }
      }]`;
      const expectedTaskContent = `[
        {
          // Check if all required prerequisites are installed and will install them if not.
          // See https://aka.ms/teamsfx-check-prerequisites-task to know the details and how to customize the args.
          "label": "Validate & install prerequisites",
          "type": "teamsfx",
          "command": "debug-check-prerequisites",
          "args": {
              "prerequisites": [
                  "nodejs", // Validate if Node.js is installed.
                  "portOccupancy" // Validate available ports to ensure those debug ones are not occupied.
              ],
              "portOccupancy": [
                  53000, // tab service port
                  7071, // backend service port
                  9229, // backend inspector port for Node.js debugger
                  3978, // bot service port
                  9239 // bot inspector port for Node.js debugger
              ]
          }
        }]`;
      const testTasks = parse(testTaskContent) as CommentArray<CommentJSONValue>;
      const oldProjectSettings = {} as any;
      const debugContext = new DebugMigrationContext(
        migrationContext,
        testTasks,
        oldProjectSettings,
        {}
      );
      await migrateTransparentPrerequisite(debugContext);
      chai.assert.equal(
        stringify(testTasks, null, 4),
        stringify(parse(expectedTaskContent), null, 4)
      );
      chai.assert.isUndefined(debugContext.appYmlConfig.deploy?.tools?.devCert?.trust);
      chai.assert.isTrue(debugContext.appYmlConfig.deploy?.tools?.dotnet);
      chai.assert.isUndefined(debugContext.appYmlConfig.deploy?.tools?.func);
    });

    it("no prerequisite task", async () => {
      const migrationContext = await mockMigrationContext(projectPath);
      const testTaskContent = `[
        {
          // Start the local tunnel service to forward public ngrok URL to local port and inspect traffic.
          // See https://aka.ms/teamsfx-local-tunnel-task to know the details and how to customize the args.
          "label": "Start local tunnel",
          "type": "teamsfx",
          "command": "debug-start-local-tunnel",
          "args": {
              "ngrokArgs": "http 3978 --log=stdout --log-format=logfmt"
          },
          "isBackground": true,
          "problemMatcher": "$teamsfx-local-tunnel-watch"
        },
        {
            // Prepare local launch information for Tab.
            // See https://aka.ms/teamsfx-debug-set-up-tab-task to know the details and how to customize the args.
            "label": "Set up tab",
            "type": "teamsfx",
            "command": "debug-set-up-tab",
            "args": {
                "baseUrl": "https://localhost:53000"
            }
        }
      ]`;
      const expectedTaskContent = testTaskContent;
      const testTasks = parse(testTaskContent) as CommentArray<CommentJSONValue>;
      const oldProjectSettings = {} as any;
      const debugContext = new DebugMigrationContext(
        migrationContext,
        testTasks,
        oldProjectSettings,
        {}
      );
      await migrateTransparentPrerequisite(debugContext);
      chai.assert.equal(
        stringify(testTasks, null, 4),
        stringify(parse(expectedTaskContent), null, 4)
      );
      chai.assert.isUndefined(debugContext.appYmlConfig.deploy?.tools?.devCert?.trust);
      chai.assert.isUndefined(debugContext.appYmlConfig.deploy?.tools?.dotnet);
      chai.assert.isUndefined(debugContext.appYmlConfig.deploy?.tools?.func);
    });

    it("empty prerequisite", async () => {
      const migrationContext = await mockMigrationContext(projectPath);
      const testTaskContent = `[
      {
        // Check if all required prerequisites are installed and will install them if not.
        // See https://aka.ms/teamsfx-check-prerequisites-task to know the details and how to customize the args.
        "label": "Validate & install prerequisites",
        "type": "teamsfx",
        "command": "debug-check-prerequisites",
        "args": {
            "prerequisites": [],
            "portOccupancy": [
                53000, // tab service port
                7071, // backend service port
                9229, // backend inspector port for Node.js debugger
                3978, // bot service port
                9239 // bot inspector port for Node.js debugger
            ]
        }
      }]`;
      const expectedTaskContent = testTaskContent;
      const testTasks = parse(testTaskContent) as CommentArray<CommentJSONValue>;
      const oldProjectSettings = {} as any;
      const debugContext = new DebugMigrationContext(
        migrationContext,
        testTasks,
        oldProjectSettings,
        {}
      );
      await migrateTransparentPrerequisite(debugContext);
      chai.assert.equal(
        stringify(testTasks, null, 4),
        stringify(parse(expectedTaskContent), null, 4)
      );
      chai.assert.isUndefined(debugContext.appYmlConfig.deploy?.tools?.devCert?.trust);
      chai.assert.isUndefined(debugContext.appYmlConfig.deploy?.tools?.dotnet);
      chai.assert.isUndefined(debugContext.appYmlConfig.deploy?.tools?.func);
    });
  });

  describe("migrateTransparentNpmInstall", () => {
    it("happy path", async () => {
      const migrationContext = await mockMigrationContext(projectPath);
      const testTaskContent = `[
        {
					"label": "Start Teams App Locally",
					"dependsOn": [
							"Validate & install prerequisites",
							"Install npm packages",
							"Start local tunnel",
							"Set up tab",
							"Set up bot",
							"Set up SSO",
							"Build & upload Teams manifest",
							"Start services"
					],
					"dependsOrder": "sequence"
			  },
        {
					"label": "Before npm install",
					"dependsOn": "Install npm packages"
			  },
        {
					// Check if all the npm packages are installed and will install them if not.
					// See https://aka.ms/teamsfx-npm-package-task to know the details and how to customize the args.
					"label": "Install npm packages",
					"type": "teamsfx",
					"command": "debug-npm-install",
					"args": {
							"projects": [
									{
											"cwd": "\${workspaceFolder}/tabs", // comment
											"npmInstallArgs": [ // comment
													"--no-audit" // comment
											]
									},
									{
											"cwd": "\${workspaceFolder}/api", // comment
											"npmInstallArgs": [] // comment
									},
									{
											"cwd": "\${workspaceFolder}/bot" // comment
									}
							]
					}
      }]`;
      const expectedTaskContent = `[
        {
					"label": "Start Teams App Locally",
					"dependsOn": [
							"Validate & install prerequisites",
							"Start local tunnel",
							"Set up tab",
							"Set up bot",
							"Set up SSO",
							"Build & upload Teams manifest",
							"Start services"
					],
					"dependsOrder": "sequence"
			  },
        {
					"label": "Before npm install"
			  }
      ]`;
      const testTasks = parse(testTaskContent) as CommentArray<CommentJSONValue>;
      const oldProjectSettings = {} as any;
      const debugContext = new DebugMigrationContext(
        migrationContext,
        testTasks,
        oldProjectSettings,
        {}
      );
      await migrateTransparentNpmInstall(debugContext);
      chai.assert.equal(
        stringify(testTasks, null, 4),
        stringify(parse(expectedTaskContent), null, 4)
      );
      chai.assert.equal(debugContext.appYmlConfig.deploy?.npmCommands?.length, 3);
      chai.assert.deepEqual(debugContext.appYmlConfig.deploy?.npmCommands?.[0], {
        args: "install --no-audit",
        workingDirectory: "./tabs",
      });
      chai.assert.deepEqual(debugContext.appYmlConfig.deploy?.npmCommands?.[1], {
        args: "install",
        workingDirectory: "./api",
      });
      chai.assert.deepEqual(debugContext.appYmlConfig.deploy?.npmCommands?.[2], {
        args: "install",
        workingDirectory: "./bot",
      });
    });

    it("one project", async () => {
      const migrationContext = await mockMigrationContext(projectPath);
      const testTaskContent = `[
        {
					"label": "Start Teams App Locally",
					"dependsOn": [
							"Validate & install prerequisites",
							"Start local tunnel",
							"Set up tab",
							"Set up bot",
							"Set up SSO",
							"Build & upload Teams manifest",
							"Start services"
					],
					"dependsOrder": "sequence"
			  },
        {
					// Check if all the npm packages are installed and will install them if not.
					// See https://aka.ms/teamsfx-npm-package-task to know the details and how to customize the args.
					"label": "Install npm packages",
					"type": "teamsfx",
					"command": "debug-npm-install",
					"args": {
							"projects": [
									{
											"cwd": "\${workspaceFolder}/tabs",
											"npmInstallArgs": "--no-audit" // comment
									}
							]
					}
        },
        {
					"label": "Before npm install",
					"dependsOn": "Install npm packages"
			  }
      ]`;
      const expectedTaskContent = `[
        {
					"label": "Start Teams App Locally",
					"dependsOn": [
							"Validate & install prerequisites",
							"Start local tunnel",
							"Set up tab",
							"Set up bot",
							"Set up SSO",
							"Build & upload Teams manifest",
							"Start services"
					],
					"dependsOrder": "sequence"
			  },
        {
					"label": "Before npm install"
			  }
      ]`;
      const testTasks = parse(testTaskContent) as CommentArray<CommentJSONValue>;
      const oldProjectSettings = {} as any;
      const debugContext = new DebugMigrationContext(
        migrationContext,
        testTasks,
        oldProjectSettings,
        {}
      );
      await migrateTransparentNpmInstall(debugContext);
      chai.assert.equal(
        stringify(testTasks, null, 4),
        stringify(parse(expectedTaskContent), null, 4)
      );
      chai.assert.equal(debugContext.appYmlConfig.deploy?.npmCommands?.length, 1);
      chai.assert.deepEqual(debugContext.appYmlConfig.deploy?.npmCommands?.[0], {
        args: "install --no-audit",
        workingDirectory: "./tabs",
      });
    });

    it("empty projects", async () => {
      const migrationContext = await mockMigrationContext(projectPath);
      const testTaskContent = `[
        {
					// Check if all the npm packages are installed and will install them if not.
					// See https://aka.ms/teamsfx-npm-package-task to know the details and how to customize the args.
					"label": "Install npm packages",
					"type": "teamsfx",
					"command": "debug-npm-install",
					"args": {
							"projects": []
					}
        },
        {
					"label": "Start Teams App Locally",
					"dependsOn": [
							"Validate & install prerequisites",
							"Install npm packages",
							"Start local tunnel",
							"Set up tab",
							"Set up bot",
							"Set up SSO",
							"Build & upload Teams manifest",
							"Start services"
					],
					"dependsOrder": "sequence"
			  },
        {
					"label": "Before npm install",
					"dependsOn": "Install npm packages"
			  }
      ]`;
      const expectedTaskContent = `[
        {
					"label": "Start Teams App Locally",
					"dependsOn": [
							"Validate & install prerequisites",
							"Start local tunnel",
							"Set up tab",
							"Set up bot",
							"Set up SSO",
							"Build & upload Teams manifest",
							"Start services"
					],
					"dependsOrder": "sequence"
			  },
        {
					"label": "Before npm install"
			  }
      ]`;
      const testTasks = parse(testTaskContent) as CommentArray<CommentJSONValue>;
      const oldProjectSettings = {} as any;
      const debugContext = new DebugMigrationContext(
        migrationContext,
        testTasks,
        oldProjectSettings,
        {}
      );
      await migrateTransparentNpmInstall(debugContext);
      chai.assert.equal(
        stringify(testTasks, null, 4),
        stringify(parse(expectedTaskContent), null, 4)
      );
      chai.assert.isUndefined(debugContext.appYmlConfig.deploy?.npmCommands);
    });

    it("no npm tasks", async () => {
      const migrationContext = await mockMigrationContext(projectPath);
      const testTaskContent = `[]`;
      const expectedTaskContent = `[]`;
      const testTasks = parse(testTaskContent) as CommentArray<CommentJSONValue>;
      const oldProjectSettings = {} as any;
      const debugContext = new DebugMigrationContext(
        migrationContext,
        testTasks,
        oldProjectSettings,
        {}
      );
      await migrateTransparentNpmInstall(debugContext);
      chai.assert.equal(
        stringify(testTasks, null, 4),
        stringify(parse(expectedTaskContent), null, 4)
      );
      chai.assert.isUndefined(debugContext.appYmlConfig.deploy?.npmCommands);
    });

    it("npmArgs not object", async () => {
      const migrationContext = await mockMigrationContext(projectPath);
      const testTaskContent = `[
        {
					"label": "Start Teams App Locally",
					"dependsOn": [
							"Validate & install prerequisites",
              "Install npm packages"
					],
					"dependsOrder": "sequence"
			  },
        {
					"label": "Install npm packages",
					"type": "teamsfx",
					"command": "debug-npm-install",
					"args": {
							"projects": [
									1
							]
					}
        }
      ]`;
      const expectedTaskContent = `[
        {
					"label": "Start Teams App Locally",
					"dependsOn": [
							"Validate & install prerequisites"
					],
					"dependsOrder": "sequence"
			  }
      ]`;
      const testTasks = parse(testTaskContent) as CommentArray<CommentJSONValue>;
      const oldProjectSettings = {} as any;
      const debugContext = new DebugMigrationContext(
        migrationContext,
        testTasks,
        oldProjectSettings,
        {}
      );
      await migrateTransparentNpmInstall(debugContext);
      chai.assert.equal(
        stringify(testTasks, null, 4),
        stringify(parse(expectedTaskContent), null, 4)
      );
      chai.assert.isUndefined(debugContext.appYmlConfig.deploy?.npmCommands);
    });

    it("cwd not string", async () => {
      const migrationContext = await mockMigrationContext(projectPath);
      const testTaskContent = `[
        {
					"label": "Start Teams App Locally",
					"dependsOn": [
							"Validate & install prerequisites",
              "Install npm packages"
					],
					"dependsOrder": "sequence"
			  },
        {
					"label": "Install npm packages",
					"type": "teamsfx",
					"command": "debug-npm-install",
					"args": {
							"projects": [
                {
                  "cwd": 1,
                  "npmInstallArgs": "--no-audit"
                }
							]
					}
        }
      ]`;
      const expectedTaskContent = `[
        {
					"label": "Start Teams App Locally",
					"dependsOn": [
							"Validate & install prerequisites"
					],
					"dependsOrder": "sequence"
			  }
      ]`;
      const testTasks = parse(testTaskContent) as CommentArray<CommentJSONValue>;
      const oldProjectSettings = {} as any;
      const debugContext = new DebugMigrationContext(
        migrationContext,
        testTasks,
        oldProjectSettings,
        {}
      );
      await migrateTransparentNpmInstall(debugContext);
      chai.assert.equal(
        stringify(testTasks, null, 4),
        stringify(parse(expectedTaskContent), null, 4)
      );
      chai.assert.isUndefined(debugContext.appYmlConfig.deploy?.npmCommands);
    });
  });

  describe("migrateTransparentLocalTunnel", () => {
    it("happy path", async () => {
      const migrationContext = await mockMigrationContext(projectPath);
      const testTaskContent = `[
        {
          // Start the local tunnel service to forward public ngrok URL to local port and inspect traffic.
          // See https://aka.ms/teamsfx-local-tunnel-task to know the details and how to customize the args.
          "label": "Start local tunnel",
          "type": "teamsfx",
          "command": "debug-start-local-tunnel",
          "args": {
              "ngrokArgs": "http 3978 --log=stdout --log-format=logfmt"
          },
          "isBackground": true,
          "problemMatcher": "$teamsfx-local-tunnel-watch"
        }
      ]`;
      const expectedTaskContent = `[
        {
          // Start the local tunnel service to forward public URL to local port and inspect traffic.
          // See https://aka.ms/teamsfx-tasks/local-tunnel for the detailed args definitions.
          "label": "Start local tunnel",
          "type": "teamsfx",
          "command": "debug-start-local-tunnel",
          "args": {
              "type": "dev-tunnel",
              "ports": [
                {
                  "portNumber": 3978,
                  "protocol": "http",
                  "access": "public",
                  "writeToEnvironmentFile": {
                    // Keep consistency with upgraded configuration.
                    "endpoint": "PROVISIONOUTPUT__AZUREWEBAPPBOTOUTPUT__SITEENDPOINT",
                    "domain": "PROVISIONOUTPUT__AZUREWEBAPPBOTOUTPUT__DOMAIN"
                  }
                }
              ],
              "env": "local",
          },
          "isBackground": true,
          "problemMatcher": "$teamsfx-local-tunnel-watch"
        }
      ]`;
      const testTasks = parse(testTaskContent) as CommentArray<CommentJSONValue>;
      const oldProjectSettings = {} as any;
      const debugContext = new DebugMigrationContext(
        migrationContext,
        testTasks,
        oldProjectSettings,
        {
          botDomain: "PROVISIONOUTPUT__AZUREWEBAPPBOTOUTPUT__DOMAIN",
          botEndpoint: "PROVISIONOUTPUT__AZUREWEBAPPBOTOUTPUT__SITEENDPOINT",
        }
      );
      await migrateTransparentLocalTunnel(debugContext);
      chai.assert.equal(
        stringify(testTasks, null, 4),
        stringify(parse(expectedTaskContent), null, 4)
      );
    });

    it("customized ngrok path", async () => {
      const migrationContext = await mockMigrationContext(projectPath);
      const testTaskContent = `[
        {
          // Start the local tunnel service to forward public ngrok URL to local port and inspect traffic.
          // See https://aka.ms/teamsfx-local-tunnel-task to know the details and how to customize the args.
          "label": "Start local tunnel",
          "type": "teamsfx",
          "command": "debug-start-local-tunnel",
          "args": {
              "ngrokArgs": "http 3978 --log=stdout --log-format=logfmt",
              "ngrokPath": "ngrok"
          },
          "isBackground": true,
          "problemMatcher": "$teamsfx-local-tunnel-watch"
        }
      ]`;
      const expectedTaskContent = `[
        {
          // Start the local tunnel service to forward public URL to local port and inspect traffic.
          // See https://aka.ms/teamsfx-tasks/local-tunnel for the detailed args definitions.
          "label": "Start local tunnel",
          "type": "teamsfx",
          "command": "debug-start-local-tunnel",
          "args": {
            "type": "dev-tunnel",
            "ports": [
              {
                "portNumber": 3978,
                "protocol": "http",
                "access": "public",
                "writeToEnvironmentFile": {
                  // Keep consistency with upgraded configuration.
                  "endpoint": "PROVISIONOUTPUT__AZUREWEBAPPBOTOUTPUT__SITEENDPOINT",
                  "domain": "PROVISIONOUTPUT__AZUREWEBAPPBOTOUTPUT__DOMAIN"
                }
              }
            ],
            "env": "local",
          },
          "isBackground": true,
          "problemMatcher": "$teamsfx-local-tunnel-watch"
        }
      ]`;
      const testTasks = parse(testTaskContent) as CommentArray<CommentJSONValue>;
      const oldProjectSettings = {} as any;
      const debugContext = new DebugMigrationContext(
        migrationContext,
        testTasks,
        oldProjectSettings,
        {
          botDomain: "PROVISIONOUTPUT__AZUREWEBAPPBOTOUTPUT__DOMAIN",
          botEndpoint: "PROVISIONOUTPUT__AZUREWEBAPPBOTOUTPUT__SITEENDPOINT",
        }
      );
      await migrateTransparentLocalTunnel(debugContext);
      chai.assert.equal(
        stringify(testTasks, null, 4),
        stringify(parse(expectedTaskContent), null, 4)
      );
    });

    it("customized ngrok port", async () => {
      const migrationContext = await mockMigrationContext(projectPath);
      const testTaskContent = `[
        {
          // Start the local tunnel service to forward public ngrok URL to local port and inspect traffic.
          // See https://aka.ms/teamsfx-local-tunnel-task to know the details and how to customize the args.
          "label": "Start local tunnel",
          "type": "teamsfx",
          "command": "debug-start-local-tunnel",
          "args": {
              "ngrokArgs": "http 3999 --log=stdout --log-format=logfmt"
          },
          "isBackground": true,
          "problemMatcher": "$teamsfx-local-tunnel-watch"
        }
      ]`;
      const expectedTaskContent = `[
        {
          // Start the local tunnel service to forward public URL to local port and inspect traffic.
          // See https://aka.ms/teamsfx-tasks/local-tunnel for the detailed args definitions.
          "label": "Start local tunnel",
          "type": "teamsfx",
          "command": "debug-start-local-tunnel",
          "args": {
              "type": "dev-tunnel",
              "ports": [
                {
                  "portNumber": 3999,
                  "protocol": "http",
                  "access": "public",
                  "writeToEnvironmentFile": {
                    // Keep consistency with upgraded configuration.
                    "endpoint": "PROVISIONOUTPUT__AZUREWEBAPPBOTOUTPUT__SITEENDPOINT",
                    "domain": "PROVISIONOUTPUT__AZUREWEBAPPBOTOUTPUT__DOMAIN"
                  }
                }
              ],
              "env": "local",
          },
          "isBackground": true,
          "problemMatcher": "$teamsfx-local-tunnel-watch"
        }
      ]`;
      const testTasks = parse(testTaskContent) as CommentArray<CommentJSONValue>;
      const oldProjectSettings = {} as any;
      const debugContext = new DebugMigrationContext(
        migrationContext,
        testTasks,
        oldProjectSettings,
        {
          botDomain: "PROVISIONOUTPUT__AZUREWEBAPPBOTOUTPUT__DOMAIN",
          botEndpoint: "PROVISIONOUTPUT__AZUREWEBAPPBOTOUTPUT__SITEENDPOINT",
        }
      );
      await migrateTransparentLocalTunnel(debugContext);
      chai.assert.equal(
        stringify(testTasks, null, 4),
        stringify(parse(expectedTaskContent), null, 4)
      );
    });

    it("customized ngrok command", async () => {
      const migrationContext = await mockMigrationContext(projectPath);
      const testTaskContent = `[
        {
          // Start the local tunnel service to forward public ngrok URL to local port and inspect traffic.
          // See https://aka.ms/teamsfx-local-tunnel-task to know the details and how to customize the args.
          "label": "Start local tunnel 1",
          "type": "teamsfx",
          "command": "debug-start-local-tunnel",
          "args": {
              "ngrokArgs": "ngrok start dev"
          },
          "isBackground": true,
          "problemMatcher": "$teamsfx-local-tunnel-watch"
        },
        {
          // Start the local tunnel service to forward public ngrok URL to local port and inspect traffic.
          // See https://aka.ms/teamsfx-local-tunnel-task to know the details and how to customize the args.
          "label": "Start local tunnel 2",
          "type": "teamsfx",
          "command": "debug-start-local-tunnel",
          "args": {
              "ngrokArgs": "http 3999 --log=stdout --log-format=logfmt"
          },
          "isBackground": true,
          "problemMatcher": "$teamsfx-local-tunnel-watch"
        }
      ]`;
      const expectedTaskContent = `[
        {
          // Microsoft 365 Agents Toolkit now uses Dev Tunnel as default tunnel solution.
          // See https://aka.ms/teamsfx-tasks/local-tunnel for more details.
          // If you still prefer to use ngrok, please refer to https://aka.ms/teamsfx-tasks/customize-tunnel-service to learn how to use your own tunnel service.
          "label": "Start local tunnel 1",
          "type": "shell",
          "command": "echo 'Microsoft 365 Agents Toolkit now uses Dev Tunnel as default tunnel solution. For manual updates, see https://aka.ms/teamsfx-tasks/local-tunnel.' && exit 1",
          "windows": {
              "options": {
                  "shell": {
                      "executable": "cmd.exe",
                      "args": [
                          "/d", "/c"
                      ]
                  }
              }
          }
        },
        {
          // Start the local tunnel service to forward public URL to local port and inspect traffic.
          // See https://aka.ms/teamsfx-tasks/local-tunnel for the detailed args definitions.
          "label": "Start local tunnel 2",
          "type": "teamsfx",
          "command": "debug-start-local-tunnel",
          "args": {
              "type": "dev-tunnel",
              "ports": [
                {
                  "portNumber": 3999,
                  "protocol": "http",
                  "access": "public",
                  "writeToEnvironmentFile": {
                    // Keep consistency with upgraded configuration.
                    "endpoint": "PROVISIONOUTPUT__AZUREWEBAPPBOTOUTPUT__SITEENDPOINT",
                    "domain": "PROVISIONOUTPUT__AZUREWEBAPPBOTOUTPUT__DOMAIN"
                  }
                }
              ],
              "env": "local",
          },
          "isBackground": true,
          "problemMatcher": "$teamsfx-local-tunnel-watch"
        }
      ]`;
      const testTasks = parse(testTaskContent) as CommentArray<CommentJSONValue>;
      const oldProjectSettings = {} as any;
      const debugContext = new DebugMigrationContext(
        migrationContext,
        testTasks,
        oldProjectSettings,
        {
          botDomain: "PROVISIONOUTPUT__AZUREWEBAPPBOTOUTPUT__DOMAIN",
          botEndpoint: "PROVISIONOUTPUT__AZUREWEBAPPBOTOUTPUT__SITEENDPOINT",
        }
      );
      await migrateTransparentLocalTunnel(debugContext);
      chai.assert.equal(
        stringify(testTasks, null, 4),
        stringify(parse(expectedTaskContent), null, 4)
      );
    });
  });

  describe("migrateSetUpTab", () => {
    it("happy path", async () => {
      const migrationContext = await mockMigrationContext(projectPath);
      const testTaskContent = `[
        {
          "label": "Start Teams App Locally",
          "dependsOn": [
              "Validate & install prerequisites",
              "Install npm packages",
              "Start local tunnel",
              "Set up tab",
              "Set up bot",
              "Set up SSO",
              "Build & upload Teams manifest",
              "Start services"
          ],
          "dependsOrder": "sequence"
        },
        {
          // Prepare local launch information for Tab.
          // See https://aka.ms/teamsfx-debug-set-up-tab-task to know the details and how to customize the args.
          "label": "Set up tab",
          "type": "teamsfx",
          "command": "debug-set-up-tab",
          "args": {
              "baseUrl": "https://localhost:53000"
          }
        }
      ]`;
      const content = `[
        {
          "label": "Start Teams App Locally",
          "dependsOn": [
              "Validate & install prerequisites",
              "Install npm packages",
              "Start local tunnel",
              "Provision",
              "Deploy",
              "Set up bot",
              "Set up SSO",
              "Build & upload Teams manifest",
              "Start services"
          ],
          "dependsOrder": "sequence"
        },
      ]`;
      const expectedTasks = parse(content) as CommentArray<CommentJSONValue>;
      expectedTasks.push(
        debugV3MigrationUtils.createResourcesTask("Provision"),
        debugV3MigrationUtils.setUpLocalProjectsTask("Deploy")
      );
      const testTasks = parse(testTaskContent) as CommentArray<CommentJSONValue>;
      const oldProjectSettings = {} as any;
      const debugContext = new DebugMigrationContext(
        migrationContext,
        testTasks,
        oldProjectSettings,
        {}
      );
      await migrateSetUpTab(debugContext);
      chai.assert.equal(stringify(debugContext.tasks, null, 4), stringify(expectedTasks, null, 4));
      chai.assert.equal(
        debugContext.appYmlConfig.provision?.configureApp?.tab?.domain,
        "localhost:53000"
      );
      chai.assert.equal(
        debugContext.appYmlConfig.provision?.configureApp?.tab?.endpoint,
        "https://localhost:53000"
      );
      chai.assert.equal(debugContext.appYmlConfig.deploy?.tab?.port, 53000);
    });
  });

  describe("migrateSetUpBot", () => {
    const botEndpointPlaceholder = "PROVISIONOUTPUT__WEBAPPOUTPUT__SITEENDPOINT";
    const placeholderMapping: debugV3MigrationUtils.DebugPlaceholderMapping = {
      botEndpoint: botEndpointPlaceholder,
    };

    let localEnvs: { [key: string]: string } = {};

    beforeEach(() => {
      sinon.stub(debugV3MigrationUtils, "updateLocalEnv").callsFake(async (context, envs) => {
        localEnvs = envs;
      });
      sinon.stub(LocalCrypto.prototype, "encrypt").callsFake((plaintext) => {
        return ok("crypto_" + plaintext);
      });
    });

    afterEach(() => {
      sinon.restore();
      localEnvs = {};
    });

    it("happy path", async () => {
      const migrationContext = await mockMigrationContext(projectPath);
      const testTaskContent = `[
        {
          "label": "Start Teams App Locally",
          "dependsOn": [
              "Validate & install prerequisites",
              "Install npm packages",
              "Start local tunnel",
              "Set up tab",
              "Set up bot",
              "Set up SSO",
              "Build & upload Teams manifest",
              "Start services"
          ],
          "dependsOrder": "sequence"
        },
        {
          // Register resources and prepare local launch information for Bot.
          // See https://aka.ms/teamsfx-debug-set-up-bot-task to know the details and how to customize the args.
          "label": "Set up bot",
          "type": "teamsfx",
          "command": "debug-set-up-bot",
          "args": {
              //// Enter your own bot information if using the existing bot. ////
              // "botId": "",
              // "botPassword": "", // use plain text or environment variable reference like $\{env:BOT_PASSWORD}
              "botMessagingEndpoint": "/api/messages" // use your own routing "/any/path", or full URL "https://contoso.com/any/path"
          }
        }
      ]`;
      const content = `[
        {
          "label": "Start Teams App Locally",
          "dependsOn": [
              "Validate & install prerequisites",
              "Install npm packages",
              "Start local tunnel",
              "Set up tab",
              "Provision",
              "Deploy",
              "Set up SSO",
              "Build & upload Teams manifest",
              "Start services"
          ],
          "dependsOrder": "sequence"
        },
      ]`;
      const expectedTasks = parse(content) as CommentArray<CommentJSONValue>;
      expectedTasks.push(
        debugV3MigrationUtils.createResourcesTask("Provision"),
        debugV3MigrationUtils.setUpLocalProjectsTask("Deploy")
      );
      const testTasks = parse(testTaskContent) as CommentArray<CommentJSONValue>;
      const oldProjectSettings = {} as any;
      const debugContext = new DebugMigrationContext(
        migrationContext,
        testTasks,
        oldProjectSettings,
        placeholderMapping
      );
      await migrateSetUpBot(debugContext);
      chai.assert.equal(stringify(debugContext.tasks, null, 4), stringify(expectedTasks, null, 4));
      chai.assert.deepEqual(localEnvs, {});
      chai.assert.equal(
        debugContext.appYmlConfig.provision?.bot?.messagingEndpoint,
        `$\{{${botEndpointPlaceholder}}}/api/messages`
      );
      chai.assert.equal(debugContext.appYmlConfig.deploy?.bot, true);
    });

    it("botMessagingEndpoint starts with http", async () => {
      const migrationContext = await mockMigrationContext(projectPath);
      const botMessagingEndpoint = "https://test.ngrok.io/api/messages";
      const testTaskContent = `[
        {
          "label": "Set up bot",
          "type": "teamsfx",
          "command": "debug-set-up-bot",
          "args": {
              "botMessagingEndpoint": "${botMessagingEndpoint}"
          }
        }
      ]`;
      const testTasks = parse(testTaskContent) as CommentArray<CommentJSONValue>;
      const oldProjectSettings = {} as any;
      const debugContext = new DebugMigrationContext(
        migrationContext,
        testTasks,
        oldProjectSettings,
        placeholderMapping
      );
      await migrateSetUpBot(debugContext);
      chai.assert.deepEqual(localEnvs, {});
      chai.assert.equal(
        debugContext.appYmlConfig.provision?.bot?.messagingEndpoint,
        botMessagingEndpoint
      );
      chai.assert.equal(debugContext.appYmlConfig.deploy?.bot, true);
    });

    it("customized botId and literal botPassword", async () => {
      const migrationContext = await mockMigrationContext(projectPath);
      const botId = "botId";
      const botPassword = "botPassword";
      const testTaskContent = `[
        {
          "label": "Set up bot",
          "type": "teamsfx",
          "command": "debug-set-up-bot",
          "args": {
              "botId": "${botId}",
              "botPassword": "${botPassword}",
          }
        }
      ]`;
      const testTasks = parse(testTaskContent) as CommentArray<CommentJSONValue>;
      const oldProjectSettings = {} as any;
      const debugContext = new DebugMigrationContext(
        migrationContext,
        testTasks,
        oldProjectSettings,
        placeholderMapping
      );
      await migrateSetUpBot(debugContext);
      chai.assert.deepEqual(localEnvs, {
        BOT_ID: botId,
        SECRET_BOT_PASSWORD: "crypto_" + botPassword,
      });
      chai.assert.equal(
        debugContext.appYmlConfig.provision?.bot?.messagingEndpoint,
        `$\{{${botEndpointPlaceholder}}}/api/messages`
      );
      chai.assert.equal(debugContext.appYmlConfig.deploy?.bot, true);
    });

    it("customized botId and env-referenced botPassword", async () => {
      const migrationContext = await mockMigrationContext(projectPath);
      const botId = "botId";
      const botPassword = "botPassword";
      process.env.BOT_PASSWORD = botPassword;
      const testTaskContent = `[
        {
          "label": "Set up bot",
          "type": "teamsfx",
          "command": "debug-set-up-bot",
          "args": {
              "botId": "${botId}",
              "botPassword": "\${env:BOT_PASSWORD}",
          }
        }
      ]`;
      const testTasks = parse(testTaskContent) as CommentArray<CommentJSONValue>;
      const oldProjectSettings = {} as any;
      const debugContext = new DebugMigrationContext(
        migrationContext,
        testTasks,
        oldProjectSettings,
        placeholderMapping
      );
      await migrateSetUpBot(debugContext);
      chai.assert.deepEqual(localEnvs, {
        BOT_ID: botId,
        SECRET_BOT_PASSWORD: "crypto_" + botPassword,
      });
      chai.assert.equal(
        debugContext.appYmlConfig.provision?.bot?.messagingEndpoint,
        `$\{{${botEndpointPlaceholder}}}/api/messages`
      );
      chai.assert.equal(debugContext.appYmlConfig.deploy?.bot, true);
      delete process.env.BOT_PASSWORD;
    });
  });

  describe("migrateSetUpSSO", () => {
    let localEnvs: { [key: string]: string } = {};

    beforeEach(() => {
      sinon.stub(debugV3MigrationUtils, "updateLocalEnv").callsFake(async (context, envs) => {
        localEnvs = envs;
      });
      sinon.stub(LocalCrypto.prototype, "encrypt").callsFake((plaintext) => {
        return ok("crypto_" + plaintext);
      });
    });

    afterEach(() => {
      sinon.restore();
      localEnvs = {};
    });

    it("happy path", async () => {
      const migrationContext = await mockMigrationContext(projectPath);
      const testTaskContent = `[
        {
          "label": "Start Teams App Locally",
          "dependsOn": [
              "Validate & install prerequisites",
              "Install npm packages",
              "Start local tunnel",
              "Set up tab",
              "Set up bot",
              "Set up SSO",
              "Build & upload Teams manifest",
              "Start services"
          ],
          "dependsOrder": "sequence"
        },
        {
          // Register resources and prepare local launch information for SSO functionality.
          // See https://aka.ms/teamsfx-debug-set-up-sso-task to know the details and how to customize the args.
          "label": "Set up SSO",
          "type": "teamsfx",
          "command": "debug-set-up-sso",
          "args": {
              //// Enter your own AAD app information if using the existing AAD app. ////
              // "objectId": "",
              // "clientId": "",
              // "clientSecret": "", // use plain text or environment variable reference like $\{env:CLIENT_SECRET}
              // "accessAsUserScopeId": "
          }
        }
      ]`;
      const content = `[
        {
          "label": "Start Teams App Locally",
          "dependsOn": [
              "Validate & install prerequisites",
              "Install npm packages",
              "Start local tunnel",
              "Set up tab",
              "Set up bot",
              "Provision",
              "Deploy",
              "Build & upload Teams manifest",
              "Start services"
          ],
          "dependsOrder": "sequence"
        },
      ]`;
      const expectedTasks = parse(content) as CommentArray<CommentJSONValue>;
      expectedTasks.push(
        debugV3MigrationUtils.createResourcesTask("Provision"),
        debugV3MigrationUtils.setUpLocalProjectsTask("Deploy")
      );
      const testTasks = parse(testTaskContent) as CommentArray<CommentJSONValue>;
      const oldProjectSettings = {} as any;
      const debugContext = new DebugMigrationContext(
        migrationContext,
        testTasks,
        oldProjectSettings,
        {}
      );
      await migrateSetUpSSO(debugContext);
      chai.assert.equal(stringify(debugContext.tasks, null, 4), stringify(expectedTasks, null, 4));
      chai.assert.equal(debugContext.appYmlConfig.provision?.registerApp?.aad, true);
      chai.assert.equal(debugContext.appYmlConfig.provision?.configureApp?.aad, true);
      chai.assert.equal(debugContext.appYmlConfig.deploy?.sso, true);
    });

    it("customized aad and literal password", async () => {
      const migrationContext = await mockMigrationContext(projectPath);
      const objectId = "objectId";
      const clientId = "clientId";
      const clientSecret = "clientSecret";
      const accessAsUserScopeId = "accessAsUserScopeId";
      const testTaskContent = `[
        {
          "label": "Set up SSO",
          "type": "teamsfx",
          "command": "debug-set-up-sso",
          "args": {
              "objectId": "${objectId}",
              "clientId": "${clientId}",
              "clientSecret": "${clientSecret}",
              "accessAsUserScopeId": "${accessAsUserScopeId}"
          }
        }
      ]`;
      const testTasks = parse(testTaskContent) as CommentArray<CommentJSONValue>;
      const oldProjectSettings = {} as any;
      const debugContext = new DebugMigrationContext(
        migrationContext,
        testTasks,
        oldProjectSettings,
        {}
      );
      await migrateSetUpSSO(debugContext);
      chai.assert.deepEqual(localEnvs, {
        AAD_APP_OBJECT_ID: objectId,
        AAD_APP_CLIENT_ID: clientId,
        SECRET_AAD_APP_CLIENT_SECRET: "crypto_" + clientSecret,
        AAD_APP_ACCESS_AS_USER_PERMISSION_ID: accessAsUserScopeId,
      });
      chai.assert.equal(debugContext.appYmlConfig.provision?.registerApp?.aad, true);
      chai.assert.equal(debugContext.appYmlConfig.provision?.configureApp?.aad, true);
      chai.assert.equal(debugContext.appYmlConfig.deploy?.sso, true);
    });

    it("customized aad and env-referenced password", async () => {
      const migrationContext = await mockMigrationContext(projectPath);
      const objectId = "objectId";
      const clientId = "clientId";
      const clientSecret = "clientSecret";
      process.env.CLIENT_SECRET = clientSecret;
      const accessAsUserScopeId = "accessAsUserScopeId";
      const testTaskContent = `[
        {
          "label": "Set up SSO",
          "type": "teamsfx",
          "command": "debug-set-up-sso",
          "args": {
              "objectId": "${objectId}",
              "clientId": "${clientId}",
              "clientSecret": "$\{env:CLIENT_SECRET}",
              "accessAsUserScopeId": "${accessAsUserScopeId}"
          }
        }
      ]`;
      const testTasks = parse(testTaskContent) as CommentArray<CommentJSONValue>;
      const oldProjectSettings = {} as any;
      const debugContext = new DebugMigrationContext(
        migrationContext,
        testTasks,
        oldProjectSettings,
        {}
      );
      await migrateSetUpSSO(debugContext);
      chai.assert.deepEqual(localEnvs, {
        AAD_APP_OBJECT_ID: objectId,
        AAD_APP_CLIENT_ID: clientId,
        SECRET_AAD_APP_CLIENT_SECRET: "crypto_" + clientSecret,
        AAD_APP_ACCESS_AS_USER_PERMISSION_ID: accessAsUserScopeId,
      });
      chai.assert.equal(debugContext.appYmlConfig.provision?.registerApp?.aad, true);
      chai.assert.equal(debugContext.appYmlConfig.provision?.configureApp?.aad, true);
      chai.assert.equal(debugContext.appYmlConfig.deploy?.sso, true);
      delete process.env.CLIENT_SECRET;
    });
  });

  describe("migratePrepareManifest", () => {
    it("without appPackagePath arg", async () => {
      const migrationContext = await mockMigrationContext(projectPath);
      const testTaskContent = `[
        {
          "label": "Start Teams App Locally",
          "dependsOn": [
              "Validate & install prerequisites",
              "Install npm packages",
              "Start local tunnel",
              "Set up tab",
              "Set up bot",
              "Set up SSO",
              "Build & upload Teams manifest",
              "Start services"
          ],
          "dependsOrder": "sequence"
        },
        {
          // Build and upload Teams manifest.
          // See https://aka.ms/teamsfx-debug-prepare-manifest-task to know the details and how to customize the args.
          "label": "Build & upload Teams manifest",
          "type": "teamsfx",
          "command": "debug-prepare-manifest",
          "args": {
              //// Enter your own Teams app package path if using the existing Teams manifest. ////
              // "appPackagePath": ""
          }
        }
      ]`;
      const content = `[
        {
          "label": "Start Teams App Locally",
          "dependsOn": [
              "Validate & install prerequisites",
              "Install npm packages",
              "Start local tunnel",
              "Set up tab",
              "Set up bot",
              "Set up SSO",
              "Provision",
              "Deploy",
              "Start services"
          ],
          "dependsOrder": "sequence"
        },
      ]`;
      const expectedTasks = parse(content) as CommentArray<CommentJSONValue>;
      expectedTasks.push(
        debugV3MigrationUtils.createResourcesTask("Provision"),
        debugV3MigrationUtils.setUpLocalProjectsTask("Deploy")
      );
      const testTasks = parse(testTaskContent) as CommentArray<CommentJSONValue>;
      const oldProjectSettings = {} as any;
      const debugContext = new DebugMigrationContext(
        migrationContext,
        testTasks,
        oldProjectSettings,
        {}
      );
      await migratePrepareManifest(debugContext);
      chai.assert.equal(stringify(debugContext.tasks, null, 4), stringify(expectedTasks, null, 4));
      chai.assert.equal(debugContext.appYmlConfig.provision?.registerApp?.teamsApp, true);
      chai.assert.equal(
        debugContext.appYmlConfig.provision?.configureApp?.teamsApp?.appPackagePath,
        undefined
      );
    });

    it("with appPackagePath", async () => {
      const migrationContext = await mockMigrationContext(projectPath);
      const testTaskContent = `[
        {
          "label": "Start Teams App Locally",
          "dependsOn": [
              "Build & upload Teams manifest",
          ],
          "dependsOrder": "sequence"
        },
        {
          "label": "Build & upload Teams manifest",
          "type": "teamsfx",
          "command": "debug-prepare-manifest",
          "args": {
              "appPackagePath": "/path/to/appPackage.zip"
          }
        }
      ]`;
      const content = `[
        {
          "label": "Start Teams App Locally",
          "dependsOn": [
              "Provision",
              "Deploy"
          ],
          "dependsOrder": "sequence"
        },
      ]`;
      const expectedTasks = parse(content) as CommentArray<CommentJSONValue>;
      expectedTasks.push(
        debugV3MigrationUtils.createResourcesTask("Provision"),
        debugV3MigrationUtils.setUpLocalProjectsTask("Deploy")
      );
      const testTasks = parse(testTaskContent) as CommentArray<CommentJSONValue>;
      const oldProjectSettings = {} as any;
      const debugContext = new DebugMigrationContext(
        migrationContext,
        testTasks,
        oldProjectSettings,
        {}
      );
      await migratePrepareManifest(debugContext);
      chai.assert.equal(stringify(debugContext.tasks, null, 4), stringify(expectedTasks, null, 4));
      chai.assert.equal(debugContext.appYmlConfig.provision?.registerApp?.teamsApp, undefined);
      chai.assert.equal(
        debugContext.appYmlConfig.provision?.configureApp?.teamsApp?.appPackagePath,
        "/path/to/appPackage.zip"
      );
    });

    it("appPackagePath not string", async () => {
      const migrationContext = await mockMigrationContext(projectPath);
      const testTaskContent = `[
        {
          "label": "Start Teams App Locally",
          "dependsOn": [
              "Build & upload Teams manifest",
          ],
          "dependsOrder": "sequence"
        },
        {
          "label": "Build & upload Teams manifest",
          "type": "teamsfx",
          "command": "debug-prepare-manifest",
          "args": {
              "appPackagePath": 1
          }
        }
      ]`;
      const content = `[
        {
          "label": "Start Teams App Locally",
          "dependsOn": [
              "Provision",
              "Deploy"
          ],
          "dependsOrder": "sequence"
        },
      ]`;
      const expectedTasks = parse(content) as CommentArray<CommentJSONValue>;
      expectedTasks.push(
        debugV3MigrationUtils.createResourcesTask("Provision"),
        debugV3MigrationUtils.setUpLocalProjectsTask("Deploy")
      );
      const testTasks = parse(testTaskContent) as CommentArray<CommentJSONValue>;
      const oldProjectSettings = {} as any;
      const debugContext = new DebugMigrationContext(
        migrationContext,
        testTasks,
        oldProjectSettings,
        {}
      );
      await migratePrepareManifest(debugContext);
      chai.assert.equal(stringify(debugContext.tasks, null, 4), stringify(expectedTasks, null, 4));
      chai.assert.equal(debugContext.appYmlConfig.provision?.registerApp?.teamsApp, true);
      chai.assert.equal(
        debugContext.appYmlConfig.provision?.configureApp?.teamsApp?.appPackagePath,
        undefined
      );
    });
  });

  describe("migrateFrontendStart", () => {
    it("happy path", async () => {
      const migrationContext = await mockMigrationContext(projectPath);
      const testTaskContent = `[
        {
          "label": "Start Frontend",
          "dependsOn": [
              "teamsfx: frontend start",
              "teamsfx: auth start"
          ],
          "dependsOrder": "parallel"
        }
      ]`;
      const content = `[
        {
          "label": "Start Frontend",
          "dependsOn": [
              "Start frontend",
              "teamsfx: auth start"
          ],
          "dependsOrder": "parallel"
        },
        {
          "label": "Start frontend",
          "type": "shell",
          "command": "npx env-cmd --silent -f .localConfigs react-scripts start",
          "isBackground": true,
          "options": {
              "cwd": "\${workspaceFolder}/tabs"
          },
          "problemMatcher": {
              "pattern": {
                  "regexp": "^.*$",
                  "file": 0,
                  "location": 1,
                  "message": 2
              },
              "background": {
                  "activeOnStart": true,
                  "beginsPattern": ".*",
                  "endsPattern": "Compiled|Failed|compiled|failed"
              }
          }
        }
      ]`;
      const testTasks = parse(testTaskContent) as CommentArray<CommentJSONValue>;
      const expectedTasks = parse(content) as CommentArray<CommentJSONValue>;
      const oldProjectSettings = {
        solutionSettings: {
          activeResourcePlugins: ["fx-resource-frontend-hosting", "fx-resource-aad-app-for-teams"],
        },
      } as any;
      const debugContext = new DebugMigrationContext(
        migrationContext,
        testTasks,
        oldProjectSettings,
        {}
      );
      await migrateFrontendStart(debugContext);
      chai.assert.equal(stringify(debugContext.tasks, null, 4), stringify(expectedTasks, null, 4));
      chai.assert.deepEqual(debugContext.appYmlConfig.deploy?.frontendStart, {
        sso: true,
        functionName: undefined,
      });
      chai.assert.equal(debugContext.appYmlConfig.deploy?.npmCommands?.length, 1);
      if (debugContext.appYmlConfig.deploy?.npmCommands) {
        chai.assert.deepEqual(debugContext.appYmlConfig.deploy?.npmCommands[0], {
          args: "install -D env-cmd",
          workingDirectory: ".",
        });
      }
    });
  });

  describe("migrateAuthStart", () => {
    afterEach(() => {
      sinon.restore();
    });

    it("happy path", async () => {
      const migrationContext = await mockMigrationContext(projectPath);
      const testTaskContent = `[
        {
          "label": "Start Frontend",
          "dependsOn": [
            "teamsfx: frontend start",
            "teamsfx: auth start"
          ],
          "dependsOrder": "parallel"
        },
      ]`;
      const testTasks = parse(testTaskContent) as CommentArray<CommentJSONValue>;
      const oldProjectSettings = {} as any;
      const debugContext = new DebugMigrationContext(
        migrationContext,
        testTasks,
        oldProjectSettings,
        {}
      );
      await migrateAuthStart(debugContext);
      chai.assert.deepEqual(debugContext.appYmlConfig.deploy?.authStart, {
        appsettingsPath: path.join(
          os.homedir(),
          ".fx",
          "localauth",
          "appsettings.Development.json"
        ),
      });
    });
  });

  describe("migrateBackendStart", () => {
    afterEach(() => {
      sinon.restore();
    });

    it("happy path", async () => {
      const migrationContext = await mockMigrationContext(projectPath);
      const testTaskContent = `[
        {
          "label": "Start Backend",
          "dependsOn": "teamsfx: backend start"
        }
      ]`;
      const testTasks = parse(testTaskContent) as CommentArray<CommentJSONValue>;
      const oldProjectSettings = {} as any;
      const debugContext = new DebugMigrationContext(
        migrationContext,
        testTasks,
        oldProjectSettings,
        {}
      );
      await migrateBackendStart(debugContext);
      chai.assert.equal(debugContext.appYmlConfig.deploy?.backendStart, true);
      chai.assert.equal(debugContext.appYmlConfig.deploy?.npmCommands?.length, 1);
      if (debugContext.appYmlConfig.deploy?.npmCommands) {
        chai.assert.deepEqual(debugContext.appYmlConfig.deploy?.npmCommands[0], {
          args: "install -D env-cmd",
          workingDirectory: ".",
        });
      }
    });
  });

  describe("migrateBackendWatch", () => {
    it("happy path", async () => {
      const migrationContext = await mockMigrationContext(projectPath);
      const testTaskContent = `[
        {
          "label": "prepare local environment",
          "type": "shell",
          "command": "echo \${command:fx-extension.pre-debug-check}"
        },
        {
          "label": "Start Backend",
          "dependsOn": [
              "teamsfx: backend watch"
          ],
          "dependsOrder": "sequence"
        }
      ]`;
      const content = `[
        {
          "label": "prepare local environment",
          "type": "shell",
          "command": "echo \${command:fx-extension.pre-debug-check}"
        },
        {
          "label": "Start Backend",
          "dependsOn": [
              "Watch backend"
          ],
          "dependsOrder": "sequence"
        },
        {
          "label": "Watch backend",
          "type": "shell",
          "command": "tsc --watch",
          "isBackground": true,
          "options": {
              "cwd": "\${workspaceFolder}/api"
          },
          "problemMatcher": "$tsc-watch",
          "presentation": {
              "reveal": "silent"
          }
        }
      ]`;
      const testTasks = parse(testTaskContent) as CommentArray<CommentJSONValue>;
      const expectedTasks = parse(content) as CommentArray<CommentJSONValue>;
      const oldProjectSettings = {} as any;
      const debugContext = new DebugMigrationContext(
        migrationContext,
        testTasks,
        oldProjectSettings,
        {}
      );
      await migrateBackendWatch(debugContext);
      chai.assert.equal(stringify(debugContext.tasks, null, 4), stringify(expectedTasks, null, 4));
    });
  });

  describe("migrateNgrokStartTask", () => {
    it("multiple ngrok label", async () => {
      const testTaskContent = `[
        {
          "label": "start ngrok",
          "dependsOn": "teamsfx: ngrok start"
        },
        {
          "label": "start ngrok 1",
          "dependsOn": ["teamsfx: ngrok start", "other label"]
        }
      ]`;
      const content = `[
        {
          "label": "start ngrok",
          "dependsOn": ["Start local tunnel"]
        },
        {
          // Start the local tunnel service to forward public URL to local port and inspect traffic.
          // See https://aka.ms/teamsfx-tasks/local-tunnel for the detailed args definitions.
          "label": "Start local tunnel",
          "type": "teamsfx",
          "command": "debug-start-local-tunnel",
          "args": {
            "type": "dev-tunnel",
            "ports": [
              {
                "portNumber": 3978,
                "protocol": "http",
                "access": "public",
                "writeToEnvironmentFile": {
                  // Keep consistency with upgraded configuration.
                  "endpoint": "PROVISIONOUTPUT__AZUREWEBAPPBOTOUTPUT__SITEENDPOINT",
                  "domain": "PROVISIONOUTPUT__AZUREWEBAPPBOTOUTPUT__DOMAIN"
                }
              }
            ],
            "env": "local",
          },
          "isBackground": true,
          "problemMatcher": "$teamsfx-local-tunnel-watch"
        },
        {
          "label": "start ngrok 1",
          "dependsOn": ["Start local tunnel", "other label"]
        }
      ]`;
      const expectedTasks = parse(content) as CommentArray<CommentJSONValue>;
      const testTasks = parse(testTaskContent) as CommentArray<CommentJSONValue>;
      const oldProjectSettings = {} as any;
      const migrationContext = await mockMigrationContext(projectPath);
      const debugContext = new DebugMigrationContext(
        migrationContext,
        testTasks,
        oldProjectSettings,
        {
          botDomain: "PROVISIONOUTPUT__AZUREWEBAPPBOTOUTPUT__DOMAIN",
          botEndpoint: "PROVISIONOUTPUT__AZUREWEBAPPBOTOUTPUT__SITEENDPOINT",
        }
      );
      migrateNgrokStartTask(debugContext);
      chai.assert.equal(stringify(debugContext.tasks, null, 4), stringify(expectedTasks, null, 4));
    });

    it("one ngrok label", async () => {
      const testTaskContent = `[
        {
          "label": "start ngrok",
          "dependsOn": "teamsfx: ngrok start"
        }
      ]`;
      const content = `[
        {
          "label": "start ngrok",
          "dependsOn": ["Start local tunnel"]
        },
        {
          // Start the local tunnel service to forward public URL to local port and inspect traffic.
          // See https://aka.ms/teamsfx-tasks/local-tunnel for the detailed args definitions.
          "label": "Start local tunnel",
          "type": "teamsfx",
          "command": "debug-start-local-tunnel",
          "args": {
            "type": "dev-tunnel",
            "ports": [
              {
                "portNumber": 3978,
                "protocol": "http",
                "access": "public",
                "writeToEnvironmentFile": {
                  // Keep consistency with upgraded configuration.
                  "endpoint": "PROVISIONOUTPUT__AZUREWEBAPPBOTOUTPUT__SITEENDPOINT",
                  "domain": "PROVISIONOUTPUT__AZUREWEBAPPBOTOUTPUT__DOMAIN"
                }
              }
            ],
            "env": "local",
          },
          "isBackground": true,
          "problemMatcher": "$teamsfx-local-tunnel-watch"
        }
      ]`;
      const expectedTasks = parse(content) as CommentArray<CommentJSONValue>;
      const testTasks = parse(testTaskContent) as CommentArray<CommentJSONValue>;
      const oldProjectSettings = {} as any;
      const migrationContext = await mockMigrationContext(projectPath);
      const debugContext = new DebugMigrationContext(
        migrationContext,
        testTasks,
        oldProjectSettings,
        {
          botDomain: "PROVISIONOUTPUT__AZUREWEBAPPBOTOUTPUT__DOMAIN",
          botEndpoint: "PROVISIONOUTPUT__AZUREWEBAPPBOTOUTPUT__SITEENDPOINT",
        }
      );
      migrateNgrokStartTask(debugContext);
      chai.assert.equal(stringify(debugContext.tasks, null, 4), stringify(expectedTasks, null, 4));
    });
  });

  describe("migrateNgrokStartCommand", () => {
    it("ngrok task with comment", async () => {
      const testTaskContent = `[
        {
          "label": "start ngrok",
          "type": "teamsfx",
          "command": "ngrok start",
          "isBackground": false,
          "dependsOn": [
              "bot npm install"
          ]
        }
      ]`;
      const content = `[
        {
          // Start the local tunnel service to forward public URL to local port and inspect traffic.
          // See https://aka.ms/teamsfx-tasks/local-tunnel for the detailed args definitions.
          "label": "start ngrok",
          "type": "teamsfx",
          "command": "debug-start-local-tunnel",
          "isBackground": true,
          "dependsOn": [
            "bot npm install"
          ],
          "args": {
            "type": "dev-tunnel",
            "ports": [
              {
                "portNumber": 3978,
                "protocol": "http",
                "access": "public",
                "writeToEnvironmentFile": {
                  // Keep consistency with upgraded configuration.
                  "endpoint": "PROVISIONOUTPUT__AZUREWEBAPPBOTOUTPUT__SITEENDPOINT",
                  "domain": "PROVISIONOUTPUT__AZUREWEBAPPBOTOUTPUT__DOMAIN"
                }
              }
            ],
            "env": "local",
          },
          "problemMatcher": "$teamsfx-local-tunnel-watch",
        }
      ]`;
      const expectedTasks = parse(content) as CommentArray<CommentJSONValue>;
      const testTasks = parse(testTaskContent) as CommentArray<CommentJSONValue>;
      const oldProjectSettings = {} as any;
      const migrationContext = await mockMigrationContext(projectPath);
      const debugContext = new DebugMigrationContext(
        migrationContext,
        testTasks,
        oldProjectSettings,
        {
          botDomain: "PROVISIONOUTPUT__AZUREWEBAPPBOTOUTPUT__DOMAIN",
          botEndpoint: "PROVISIONOUTPUT__AZUREWEBAPPBOTOUTPUT__SITEENDPOINT",
        }
      );
      migrateNgrokStartCommand(debugContext);
      chai.assert.equal(stringify(debugContext.tasks, null, 4), stringify(expectedTasks, null, 4));
    });
  });

  describe("migrateGetFuncPathCommand", () => {
    it("happy path", async () => {
      const migrationContext = await mockMigrationContext(projectPath);
      const testTaskContent = `[
        {
            "label": "Start bot",
            "type": "shell",
            "command": "npm run dev:teamsfx",
            "isBackground": true,
            "options": {
                "cwd": "\${workspaceFolder}/bot",
                "env": {
                    "PATH": "\${command:fx-extension.get-func-path}\${env:PATH}"
                }
            },
            "problemMatcher": {
                "pattern": {
                    "regexp": "^.*$",
                    "file": 0,
                    "location": 1,
                    "message": 2
                },
                "background": {
                    "activeOnStart": true,
                    "beginsPattern": "^.*(Job host stopped|signaling restart).*$",
                    "endsPattern": "^.*(Worker process started and initialized|Host lock lease acquired by instance ID).*$"
                }
            },
            "dependsOn": [
                "Start Azurite emulator",
                "Watch bot"
            ]
        },
        {
          "label": "Customized path",
          "type": "shell",
          "command": "npm run dev:teamsfx",
          "isBackground": true,
          "options": {
              "cwd": "\${workspaceFolder}/bot",
              "env": {
                  "PATH": "\${command:fx-extension.get-func-path}\${env:PATH}"
              }
          },
          "windows": {
            "options": {
              "env": {
                  "PATH": "\${env:PATH}\${command:fx-extension.get-func-path}"
              }
            }
          },
          "linux": {
            "options": {
              "env": {
                  "PATH": "\${command:fx-extension.get-func-path}"
              }
            }
          },
          "osx": {
            "options": {
              "env": {
                  "path": "\${command:fx-extension.get-func-path}\${env:path}\${command:fx-extension.get-func-path}"
              }
            }
          }
        }
      ]`;
      const expectedTaskContent = `[
        {
          "label": "Start bot",
          "type": "shell",
          "command": "npm run dev:teamsfx",
          "isBackground": true,
          "options": {
              "cwd": "\${workspaceFolder}/bot",
              "env": {
                  "PATH": "\${workspaceFolder}/devTools/func\${command:fx-extension.get-path-delimiter}\${env:PATH}"
              }
          },
          "problemMatcher": {
              "pattern": {
                  "regexp": "^.*$",
                  "file": 0,
                  "location": 1,
                  "message": 2
              },
              "background": {
                  "activeOnStart": true,
                  "beginsPattern": "^.*(Job host stopped|signaling restart).*$",
                  "endsPattern": "^.*(Worker process started and initialized|Host lock lease acquired by instance ID).*$"
              }
          },
          "dependsOn": [
              "Start Azurite emulator",
              "Watch bot"
          ]
      },
      {
        "label": "Customized path",
        "type": "shell",
        "command": "npm run dev:teamsfx",
        "isBackground": true,
        "options": {
            "cwd": "\${workspaceFolder}/bot",
            "env": {
                "PATH": "\${workspaceFolder}/devTools/func\${command:fx-extension.get-path-delimiter}\${env:PATH}"
            }
        },
        "windows": {
          "options": {
            "env": {
                "PATH": "\${env:PATH}\${command:fx-extension.get-path-delimiter}\${workspaceFolder}/devTools/func\${command:fx-extension.get-path-delimiter}"
            }
          }
        },
        "linux": {
          "options": {
            "env": {
                "PATH": "\${workspaceFolder}/devTools/func\${command:fx-extension.get-path-delimiter}"
            }
          }
        },
        "osx": {
          "options": {
            "env": {
                "path": "\${workspaceFolder}/devTools/func\${command:fx-extension.get-path-delimiter}\${env:path}\${command:fx-extension.get-path-delimiter}\${workspaceFolder}/devTools/func\${command:fx-extension.get-path-delimiter}"
            }
          }
        }
      }
      ]`;
      const testTasks = parse(testTaskContent) as CommentArray<CommentJSONValue>;
      const oldProjectSettings = {} as any;
      const debugContext = new DebugMigrationContext(
        migrationContext,
        testTasks,
        oldProjectSettings,
        {}
      );
      await migrateGetFuncPathCommand(debugContext);
      chai.assert.equal(
        stringify(testTasks, null, 4),
        stringify(parse(expectedTaskContent), null, 4)
      );
    });
  });
});
