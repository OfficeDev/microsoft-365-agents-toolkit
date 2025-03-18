import { err, FxError, LogProvider, ok, Result } from "@microsoft/teamsfx-api";
import { assert } from "chai";
import "mocha";
import sinon from "sinon";
import fs from "fs-extra";
import {
  DriverInstance,
  ExecutionResult,
  ProjectModel,
} from "../../../src/component/configManager/interface";
import { DriverContext } from "../../../src/component/driver/interface/commonArgs";
import { setTools } from "../../../src/common/globalVars";
import { MockTools } from "../../core/utils";
import { ExecutionResult as DriverResult } from "../../../src/component/driver/interface/stepDriver";
import {
  ProjectTypeProps,
  TelemetryProperty,
  WebApplicationIdValue,
} from "../../../src/common/telemetry";
import { metadataDAPropertiesUtil } from "../../../src/component/utils/metadataDAProperties";
import { manifestUtils } from "../../../src/component/driver/teamsApp/utils/ManifestUtils";

function mockedResolveDriverInstances(log: LogProvider): Result<DriverInstance[], FxError> {
  return ok([
    {
      uses: "arm/deploy",
      with: undefined,
      instance: {
        execute: async (args: unknown, context: DriverContext): Promise<DriverResult> => {
          return { result: ok(new Map<string, string>()), summaries: [] };
        },
      },
    },
  ]);
}

describe("metadata rsc permission util", () => {
  const manifestContent = `
  
  `;
  const sandbox = sinon.createSandbox();

  const readAppManifestRes = {
    $schema: "https://developer.microsoft.com/json-schemas/teams/v1.19/MicrosoftTeams.schema.json",
    manifestVersion: "1.19",
    version: "1.0.0",
    id: "TEAMS_APP_ID",
    developer: {
      name: "Teams App, Inc.",
      websiteUrl: "https://www.example.com",
      privacyUrl: "https://www.example.com/privacy",
      termsOfUseUrl: "https://www.example.com/termofuse",
    },
    icons: {
      color: "color.png",
      outline: "outline.png",
    },
    name: {
      short: "test_NAME_SUFFIX",
      full: "Full name for ddfdfdddd",
    },
    description: {
      short: "Short description for test",
      full: "Full description for test",
    },
    accentColor: "#FFFFFF",
    copilotAgents: {
      declarativeAgents: [
        {
          id: "declarativeAgent",
          file: "declarativeAgent.json",
        },
      ],
    },
    permissions: ["identity", "messageTeamMembers"],
    validDomains: [],
  };

  const mockProjectModel: ProjectModel = {
    version: "1.0.0",
    provision: {
      name: "provision",
      driverDefs: [
        {
          uses: "teamsApp/validateManifest",
          with: {
            manifestPath: "./appPackage/manifest.json",
          },
        },
      ],
      resolvePlaceholders: () => {
        return ["AZURE_SUBSCRIPTION_ID", "AZURE_RESOURCE_GROUP_NAME"];
      },
      execute: async (ctx: DriverContext): Promise<ExecutionResult> => {
        return { result: ok(new Map()), summaries: [] };
      },
      resolveDriverInstances: mockedResolveDriverInstances,
    },
    environmentFolderPath: "./envs",
  };
  let tools: MockTools;
  const ymlPath = "teamsapp.yml";

  beforeEach(() => {
    tools = new MockTools();
    setTools(tools);
  });

  afterEach(() => {
    sandbox.restore();
  });

  it("parseManifest happy path", async () => {
    sandbox.stub(fs, "pathExists").resolves(true);
    sandbox.stub(manifestUtils, "_readAppManifest").resolves(ok(readAppManifestRes as any));

    sandbox.stub(fs, "readJSON").callsFake(async (path: string) => {
      if (path.endsWith("declarativeAgent.json")) {
        return {
          capabilities: [
            {
              name: "capability1",
            },
            {
              name: "capability2",
            },
          ],
          actions: [
            {
              file: "action1.json",
            },
            {
              file: "action2.json",
            },
          ],
        };
      } else if (path.endsWith("action1.json")) {
        return {
          runtimes: [
            {
              type: "OpenApi",
              auth: {
                type: "None",
              },
              spec: {
                url: "apiSpecificationFile/openapi.yaml",
              },
              run_for_functions: ["deletePet"],
            },
          ],
        };
      } else if (path.endsWith("action2.json")) {
        return {
          runtimes: [
            {
              type: "OpenApi",
              auth: {
                type: "ApiKeyPluginVault",
              },
              spec: {
                url: "apiSpecificationFile/openapi.yaml",
              },
              run_for_functions: ["deletePet"],
            },
          ],
        };
      }
    });

    const props: any = {};
    await metadataDAPropertiesUtil.parseManifest(ymlPath, mockProjectModel, props);

    assert(props[ProjectTypeProps.DeclarativeAgentCapabilitiesCount] === "2");
    assert(props[ProjectTypeProps.DeclarativeAgentActionsCount] === "2");
    assert(props[ProjectTypeProps.DeclarativeAgentCapabilities] === "capability1,capability2");
    assert(props[ProjectTypeProps.DeclarativeAgentPluginAuthTypes] === "None;ApiKeyPluginVault");
  });

  it("parseManifest no manifest", async () => {
    sandbox.stub(fs, "pathExists").resolves(false);
    const props: any = {};
    await metadataDAPropertiesUtil.parseManifest(ymlPath, mockProjectModel, props);
    assert(props[ProjectTypeProps.TeamsManifestVersion] === undefined);
  });
});
