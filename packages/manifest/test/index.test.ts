import "mocha";
import * as chai from "chai";
import * as path from "path";
import fs from "fs-extra";
import chaiAsPromised from "chai-as-promised";
import sinon from "sinon";
import { jsonToManifest, ManifestUtil, TeamsAppManifest, TeamsAppManifestJSONSchema } from "../src";
chai.use(chaiAsPromised);

describe("Manifest manipulation", async () => {
  describe("loadFromPath", async () => {
    it("should succeed when loading from a valid path", async () => {
      const filePath = path.join(__dirname, "manifest.json");
      const manifest = await ManifestUtil.loadFromPath(filePath);
      chai.expect(manifest.id).equals("aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee");
    });

    it("should throw when loading from an invalid path", async () => {
      const invalidPath = path.join(__dirname, "invalid.json");
      chai.expect(await fs.pathExists(invalidPath)).equals(false);

      chai.expect(ManifestUtil.loadFromPath(invalidPath)).to.be.rejectedWith(Error);
    });
  });

  describe("writeToPath", async () => {
    const mocker = sinon.createSandbox();
    const fileContent: Map<string, string> = new Map();

    before(() => {
      // eslint-disable-next-line @typescript-eslint/ban-ts-comment
      // @ts-ignore
      mocker.stub(fs, "writeJson").callsFake((file: string, obj: any) => {
        fileContent.set(file, JSON.stringify(obj));
      });
    });

    after(() => {
      mocker.restore();
      fileContent.clear();
    });

    it("should succeed when writing to a valid path", async () => {
      const filePath = path.join(__dirname, "test_manifest.json");
      const manifest = new TeamsAppManifest();
      const fakeId = "some-fake-id";
      manifest.id = fakeId;
      await ManifestUtil.writeToPath(filePath, manifest);
      chai.expect(fileContent.get(filePath)).is.not.undefined;
      chai.expect(JSON.parse(fileContent.get(filePath)!).id).equals(fakeId);
    });
  });

  describe("validateManifest", async () => {
    const mocker = sinon.createSandbox();

    const schema = await loadSchema();

    before(() => {
      mocker.stub(ManifestUtil, "fetchSchema").resolves(schema);
    });

    after(() => {
      mocker.restore();
    });

    it("should throw if $schema is undefiend", async () => {
      const manifest = new TeamsAppManifest();
      manifest.$schema = undefined;
      chai.expect(ManifestUtil.validateManifest(manifest)).to.be.rejectedWith(Error);
    });

    it("should return empty arry when validation passes", async () => {
      const filePath = path.join(__dirname, "manifest.json");
      const validManifest = await ManifestUtil.loadFromPath(filePath);
      const result = await ManifestUtil.validateManifest(validManifest);
      chai.expect(result).to.be.empty;
    });
  });

  describe("validateManifestAgainstSchema", async () => {
    it("should return empty array when validation passes", async () => {
      const schema = await loadSchema();
      const filePath = path.join(__dirname, "manifest.json");
      const validManifest = await ManifestUtil.loadFromPath(filePath);
      const result = await ManifestUtil.validateManifestAgainstSchema(validManifest, schema);
      chai.expect(result).to.be.empty;
    });

    it("should return error string array", async () => {
      // schema has version 1.11
      const schema = await loadSchema();
      const manifest = new TeamsAppManifest();
      chai.expect(manifest.manifestVersion).equals("1.15");
      const result = await ManifestUtil.validateManifestAgainstSchema(manifest, schema);
      chai.expect(result).not.to.be.empty;
      chai.expect(result.length).equals(2);
      // 1.15 doesn't match 1.11, so it should return an error
      chai.expect(result[0]).to.contain("/manifestVersion");
    });
  });
  describe("useCopilotExtensionsInSchema", async () => {
    let fetchSchemaStub: sinon.SinonStub;

    beforeEach(() => {
      fetchSchemaStub = sinon.stub(ManifestUtil, "fetchSchema");
    });

    afterEach(() => {
      sinon.restore();
    });

    it("should return true when copilotExtensions exist in schema definitions", async () => {
      const mockSchema = {
        properties: {
          copilotExtensions: {},
        },
      };

      fetchSchemaStub.resolves(mockSchema);

      const result = await ManifestUtil.useCopilotExtensionsInSchema({} as any);
      chai.assert.isTrue(result);
    });

    it("should return false when copilotExtensions do not exist in schema definitions", async () => {
      const mockSchema = {
        properties: {},
      };
      fetchSchemaStub.resolves(mockSchema);

      const result = await ManifestUtil.useCopilotExtensionsInSchema({} as any);
      chai.assert.isFalse(result);
    });
  });
});

describe("ManifestUtil", () => {
  it("should return the correct manifest version", () => {
    const json = {
      $schema:
        "https://developer.microsoft.com/en-us/json-schemas/teams/v1.19/MicrosoftTeams.schema.json",
      manifestVersion: "1.19",
      version: 1.0,
      id: "${{TEAMS_APP_ID}}",
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
        short: "huajiecea040906${{APP_NAME_SUFFIX}}",
        full: "full name for huajiecea040906",
      },
      description: {
        short: "Repair Service",
        full: "A simple service to manage repairs",
      },
      accentColor: "#FFFFFF",
      bots: [
        {
          botId: "${{BOT_ID}}",
          scopes: ["personal", "team", "groupChat"],
          supportsFiles: false,
          isNotificationOnly: false,
          commandLists: [
            {
              scopes: ["personal"],
              commands: [
                {
                  title: "List all repairs without auth",
                  description: "List all repairs without auth",
                },
              ],
            },
          ],
        },
      ],
      composeExtensions: [],
      configurableTabs: [],
      staticTabs: [],
      permissions: ["identity", "messageTeamMembers"],
      validDomains: [],
    };

    try {
      jsonToManifest(JSON.stringify(json));
      chai.assert.fail("Expected error not thrown");
    } catch (error: any) {
      chai.assert.include(error.message, `Invalid value for key "version"`);
    }
  });
});

async function loadSchema(): Promise<any> {
  const schemaPath = path.join(__dirname, "MicrosoftTeams.schema.json");
  return fs.readJson(schemaPath);
}
