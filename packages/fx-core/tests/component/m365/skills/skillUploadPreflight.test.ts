// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { err, TeamsAppManifest, UserError } from "@microsoft/teamsfx-api";
import AdmZip from "adm-zip";
import axios from "axios";
import { expect, vi } from "vitest";
import { prepareIncrementalSkillUpload } from "../../../../src/component/m365/skills/incrementalSkillUpload";
import { MockedM365Provider } from "../../../core/utils";

function packageBytes(conflictingReference = false, invalidMetadata = false): Buffer {
  const manifest = new TeamsAppManifest();
  manifest.id = "11111111-1111-4111-8111-111111111111";
  manifest.version = "1.0.0";
  manifest.copilotAgents = {
    declarativeAgents: [{ id: "agent", file: "agent.json" }],
  };
  const zip = new AdmZip();
  zip.addFile("manifest.json", Buffer.from(JSON.stringify(manifest)));
  zip.addFile(
    "agent.json",
    Buffer.from(
      JSON.stringify({
        version: "v1.6",
        name: "Agent",
        description: "Test agent",
        instructions: "Use skills.",
        agent_skills: [{ folder: "skills/reports" }],
        ...(conflictingReference && {
          actions: [{ id: "keep", file: "skills/reports/data.json" }],
        }),
      })
    )
  );
  zip.addFile(
    "skills/reports/SKILL.md",
    Buffer.from("---\nname: reports\ndescription: Report skill\n---\nRead the data.")
  );
  zip.addFile("skills/reports/data.json", Buffer.from("{}"));
  if (invalidMetadata) {
    zip.addFile("packageMetadata.json", Buffer.from('{"builderInfo":[]}'));
  }
  return zip.toBuffer();
}

describe("incremental upload preflight", () => {
  afterEach(() => vi.restoreAllMocks());

  it.each([
    ["a reference into removed skill content", true, false],
    ["invalid transport metadata", false, true],
  ])("AC-12: rejects %s before authentication or mutation", async (_, reference, metadata) => {
    const provider = new MockedM365Provider();
    const token = vi.spyOn(provider, "getAccessToken");
    const network = vi.spyOn(axios, "create").mockImplementation(() => {
      throw new Error("Unexpected network client before preflight");
    });
    const result = await prepareIncrementalSkillUpload(
      packageBytes(reference === true, metadata === true),
      "https://titles.prod.mos.microsoft.com",
      "mos-token",
      provider
    );
    expect(result.isErr()).toBe(true);
    expect(token).not.toHaveBeenCalled();
    expect(network).not.toHaveBeenCalled();
  });

  it("AC-10: preserves an authentication failure after valid preflight", async () => {
    const provider = new MockedM365Provider();
    const failure = new UserError("test", "ConsentRequired", "Test consent required");
    const token = vi.spyOn(provider, "getAccessToken").mockResolvedValue(err(failure));
    const network = vi.spyOn(axios, "create").mockImplementation(() => {
      throw new Error("Unexpected network client after failed authentication");
    });
    const result = await prepareIncrementalSkillUpload(
      packageBytes(),
      "https://titles.prod.mos.microsoft.com",
      "mos-token",
      provider
    );
    expect(result.isErr() && result.error).toBe(failure);
    expect(token).toHaveBeenCalledTimes(1);
    expect(token).toHaveBeenCalledWith({
      scopes: ["https://graph.microsoft.com/Files.ReadWrite"],
    });
    expect(network).not.toHaveBeenCalled();
  });
});
