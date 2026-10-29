// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { Inputs, Platform } from "@microsoft/teamsfx-api";
import { describe, expect, it } from "vitest";
import {
  getTabWebsiteOptions,
  selectBotIdsQuestion,
  selectTabsContentUrlQuestion,
} from "../../src/question/create";
import { QuestionNames } from "../../src/question/questionNames";
import { createFromTdpNode } from "../../src/question/scaffold/vsc/createFromTdpNode";

describe("Developer Portal draft capability questions", () => {
  it("AC-05: tab choices include only named tabs with a corresponding URL", async () => {
    const inputs: Inputs = {
      platform: Platform.VSCode,
      teamsAppFromTdp: {
        appId: "11111111-1111-4111-8111-111111111111",
        appProfile: {
          appDetails: {
            applicationManifest: {
              staticTabs: [
                {
                  name: "Both",
                  contentUrl: "https://example.com/content",
                  websiteUrl: "https://example.com/site",
                },
                { name: "Content", contentUrl: "https://example.com/content" },
                { name: "Website", websiteUrl: "https://example.com/site" },
                {
                  name: null,
                  contentUrl: "https://example.com/content",
                  websiteUrl: "https://example.com/site",
                },
                { name: "No URLs" },
              ],
            },
          },
        },
      },
    };
    expect(getTabWebsiteOptions(inputs).map((item) => item.id)).toEqual(["Both", "Website"]);
    const options = selectTabsContentUrlQuestion().dynamicOptions;
    expect(typeof options).toBe("function");
    if (typeof options !== "function") throw new Error("Missing content URL options");
    expect(await options(inputs)).toEqual([
      expect.objectContaining({ id: "Both" }),
      expect.objectContaining({ id: "Content" }),
    ]);
    expect(await options({ platform: Platform.VSCode })).toEqual([]);
    expect(getTabWebsiteOptions({ platform: Platform.VSCode })).toEqual([]);
  });

  it.each([false, true])(
    "AC-05: bot choices support legacy and native drafts (native=%s)",
    async (native) => {
      const inputs: Inputs = {
        platform: Platform.VSCode,
        teamsAppFromTdp: native
          ? {
              appId: "resource",
              appProfile: {
                appDetails: {
                  applicationManifest: {
                    bots: [{ botId: "bot" }],
                    composeExtensions: [{ botId: "extension" }],
                  },
                },
              },
            }
          : {
              teamsAppId: "resource",
              bots: [{ botId: "bot" }],
              messagingExtensions: [{ botId: "extension" }],
            },
      };
      const options = selectBotIdsQuestion().dynamicOptions;
      expect(typeof options).toBe("function");
      if (typeof options !== "function") throw new Error("Missing bot options");
      expect(await options(inputs)).toEqual([
        expect.objectContaining({ id: "bot" }),
        expect.objectContaining({ id: "messageExtension" }),
      ]);
      expect(await options({ platform: Platform.VSCode })).toEqual([]);
      expect(
        await options({
          platform: Platform.VSCode,
          teamsAppFromTdp: {
            appId: "resource",
            appProfile: {
              appDetails: { applicationManifest: { bots: [], composeExtensions: [] } },
            },
          },
        })
      ).toEqual([]);
    }
  );

  it("AC-05: handoff asks URL replacement only when that URL exists and bot code is unnecessary", async () => {
    const group = createFromTdpNode().children?.find(
      (child) => child.data.name === QuestionNames.RepalceTabUrl
    );
    expect(group?.children).toHaveLength(2);
    for (const [websiteUrl, contentUrl, bot] of [
      [undefined, undefined, false],
      ["https://example.com/site", undefined, false],
      [undefined, "https://example.com/content", false],
      ["https://example.com/site", "https://example.com/content", true],
    ]) {
      const inputs: Inputs = {
        platform: Platform.VSCode,
        teamsAppFromTdp: {
          appId: "resource",
          appProfile: {
            appDetails: {
              applicationManifest: {
                staticTabs: [
                  {
                    name: "Tab",
                    websiteUrl: typeof websiteUrl === "string" ? websiteUrl : undefined,
                    contentUrl: typeof contentUrl === "string" ? contentUrl : undefined,
                  },
                ],
                bots: bot ? [{ botId: "bot" }] : [],
              },
            },
          },
        },
      };
      for (const node of group?.children ?? []) {
        if (typeof node.condition !== "function") throw new Error("Missing tab condition");
        const hasUrl =
          node.data.name === QuestionNames.ReplaceWebsiteUrl ? !!websiteUrl : !!contentUrl;
        expect(await node.condition(inputs)).toBe(hasUrl && !bot);
        expect(await node.condition({ platform: Platform.VSCode })).toBe(false);
        expect(
          await node.condition({
            platform: Platform.VSCode,
            teamsAppFromTdp: { appId: "resource", appProfile: null },
          })
        ).toBe(false);
      }
    }
  });
});
