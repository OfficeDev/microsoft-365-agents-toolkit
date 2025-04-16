// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { FxError, Result, TeamsManifest } from "@microsoft/teamsfx-api";
import { createHash } from "crypto";
import * as path from "path";
import { TOOLS } from "../../common/globalVars";
import { TelemetryEvent, TelemetryProperty } from "../../common/telemetry";
import { LifecycleNames, ProjectModel } from "../configManager/interface";
import { yamlParser } from "../configManager/parser";
import { metadataDAPropertiesUtil } from "./metadataDAProperties";
import { metadataGraphPermissionUtil } from "./metadataGraphPermssion";
import { metadataRscPermissionUtil } from "./metadataRscPermission";

class MetadataUtil {
  async parse(ymlPath: string): Promise<Result<ProjectModel, FxError>> {
    const res = await yamlParser.parse(ymlPath, true);
    const fileName = path.basename(ymlPath);
    const props: { [key: string]: string } = {};
    props[TelemetryProperty.YmlName] = fileName.split(".").join("");
    if (res.isOk()) {
      for (const name of LifecycleNames) {
        const str = res.value[name]?.driverDefs
          .map((def) => def.uses)
          .toString()
          .split("/")
          .join("");
        props[name + ".actions"] = str ?? "";
      }
      props[TelemetryProperty.YmlSchemaVersion] = res.value.version;
      props[TelemetryProperty.SampleAppName] = MetadataUtil.parseSampleTag(
        res.value.additionalMetadata
      );
      await metadataGraphPermissionUtil.parseAadManifest(ymlPath, res.value, props);
      await metadataRscPermissionUtil.parseManifest(ymlPath, res.value, props);
      await metadataDAPropertiesUtil.parseManifest(ymlPath, res.value, props);

      TOOLS.telemetryReporter?.sendTelemetryEvent(TelemetryEvent.MetaData, props);
    }

    return res;
  }

  static parseSampleTag(additionalMetadata: { [key: string]: unknown } | undefined): string {
    if (additionalMetadata === undefined) {
      return "";
    }

    const sampleTag = additionalMetadata["sampleTag"];
    if (typeof sampleTag === "string") {
      // replace characters that could make the tag be mistaken as a file path or an email address
      return sampleTag.replace(/[@\/\\\.]/g, "_");
    } else {
      return "";
    }
  }

  parseManifest(manifest: TeamsManifest): void {
    const props: { [key: string]: string } = {};
    const prefix = "manifest.";
    props[prefix + "id"] = manifest.id ?? "";
    props[prefix + "version"] = manifest.version ?? "";
    props[prefix + "manifestVersion"] = manifest.manifestVersion ?? "";
    props[prefix + "bots"] = manifest.bots?.map((bot) => bot.botId).toString() ?? "";
    props[prefix + "composeExtensions"] =
      manifest.composeExtensions?.map((bot) => bot.botId).toString() ?? "";
    props[prefix + "staticTabs.contentUrl"] =
      manifest.staticTabs
        ?.map((tab) =>
          tab.contentUrl ? createHash("sha256").update(tab.contentUrl).digest("hex") : "undefined"
        )
        .toString() ?? "";
    props[prefix + "configurableTabs.configurationUrl"] =
      manifest.configurableTabs
        ?.map((tab) =>
          tab.configurationUrl
            ? createHash("sha256").update(tab.configurationUrl).digest("hex")
            : "undefined"
        )
        .toString() ?? "";
    props[prefix + "webApplicationInfo.id"] = (manifest as any).webApplicationInfo?.id ?? "";
    props[prefix + "extensions"] =
      "extensions" in manifest && manifest["extensions"]?.length != 0 ? "true" : "false";

    TOOLS.telemetryReporter?.sendTelemetryEvent(TelemetryEvent.MetaData, props);
  }
}

export const metadataUtil = new MetadataUtil();
