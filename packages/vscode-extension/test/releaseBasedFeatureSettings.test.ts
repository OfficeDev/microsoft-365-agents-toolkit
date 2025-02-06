import * as chai from "chai";
import { releaseControlledFeatureSettings } from "../src/releaseBasedFeatureSettings";

describe("releaseControlledFeatureSettings", () => {
  it("verify default values", async () => {
    const settings = releaseControlledFeatureSettings;
    chai.assert.isTrue(settings.shouldEnableTeamsCopilotChatUI);
  });
});
