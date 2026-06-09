// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { FxError } from "@microsoft/teamsfx-api";
import { Result } from "neverthrow";
import templateConfig from "../../common/templates-config.json";
import { loadBundledFloor } from "../../v4/distribution/bundledFloor";
import { TemplateSource, resolveTemplateSource } from "../../v4/distribution/templateSource";
import { createTemplateSourcePort } from "../../v4/distribution/templateSourcePort";
import { defaultTryLimits } from "./constant";

/**
 * Resolve which v4 template release the metadata should come from, using the
 * SAME single decision point as the template package:
 * `resolveTemplateSource((v4.range, v4.bundled, port))`.
 *
 * `bundled` is CD-baked (= !goproduct), so goproduct builds (stable AND
 * prerelease) resolve to an `online`/`cache` source while non-goproduct/daily
 * builds resolve to the bundled floor. The metadata.zip asset lives in the same
 * `templates-v4@<version>` release as the resolved package, so the resolved
 * version names the release that `fetchOnlineTemplateMetadata` pulls metadata
 * from. Resolving here also downloads, verifies, and caches the template
 * package, so a later content scaffold reuses it with no re-download.
 *
 * An unreachable channel resolves to a bundled-fallback origin (not an error);
 * only a malformed tag list or a digest mismatch surfaces as `Result.err`.
 *
 * Transitional: remove once selector.json drives metadata distribution.
 */
export function resolveV4MetadataSource(): Promise<Result<TemplateSource, FxError>> {
  const port = createTemplateSourcePort(
    {
      templatesV4TagListURL: templateConfig.templatesV4TagListURL,
      templateDownloadBaseURL: templateConfig.templateDownloadBaseURL,
      tryLimits: defaultTryLimits,
    },
    loadBundledFloor()
  );
  return resolveTemplateSource({
    range: templateConfig.v4.range,
    bundled: templateConfig.v4.bundled,
    port,
  });
}
