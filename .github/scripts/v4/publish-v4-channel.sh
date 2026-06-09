#!/usr/bin/env bash
# Copyright (c) Microsoft Corporation.
# Licensed under the MIT license.
#
# CD wrapper (v4-isolated): publish one immutable v4 template release and append
# its digest entry to the v4 channel's NDJSON tag-list. This collapses the five
# release-action steps that previously lived inline in cd.yml into one step so
# the v4 footprint in the workflow stays small while the single-coordination
# point (the lerna-minted templates version) is preserved.
#
# It only ever touches the two v4-owned GitHub releases (`templates-v4@<ver>`
# and `template-v4-tag-list`); the v3 `templates@` channel is never read or
# written here.
#
# Usage (run AFTER the templates build, with build/v4/templates.zip present):
#   bash .github/scripts/v4/publish-v4-channel.sh \
#     <templates@<ver> tag> <path to templates.zip> <temp dir> <commit sha> \
#     [path to metadata.zip]
#
# The optional 5th argument is transitional: it mirrors the v3 `metadata.zip`
# onto the v4 release so `fetchOnlineTemplateMetadata` can pull metadata from
# the v4 tag when TEAMSFX_V4_ENABLED is on. Remove once selector.json drives
# metadata distribution.
#
# Requires: gh CLI authenticated via GH_TOKEN, node on PATH.
set -euo pipefail

TEMPLATE_TAG="${1:?Need the templates@<ver> tag (steps.version-change.outputs.TEMPLATE_VERSION).}"
ZIP="${2:?Need the path to templates.zip.}"
TMP="${3:?Need a temp directory.}"
SHA="${4:?Need the commit sha to anchor the release.}"
METADATA_ZIP="${5:-}"

VERSION="${TEMPLATE_TAG#templates@}"
TAG="templates-v4@$VERSION"
NDJSON="$TMP/template-v4-tags.ndjson"

# 1. Upsert the immutable per-version release and (re)upload the package.
if ! gh release view "$TAG" >/dev/null 2>&1; then
  gh release create "$TAG" \
    --title "$TAG" \
    --notes "v4 template package for $TAG" \
    --target "$SHA"
fi
gh release upload "$TAG" "$ZIP" --clobber

# 1b. Transitional: mirror the v3 metadata.zip onto the v4 release so the
#     engine can fetch metadata from the v4 tag when TEAMSFX_V4_ENABLED is on.
if [ -n "$METADATA_ZIP" ]; then
  gh release upload "$TAG" "$METADATA_ZIP" --clobber
fi

# 2. Pull the existing NDJSON tag-list (absent on the first release).
gh release download template-v4-tag-list --pattern template-v4-tags.ndjson --dir "$TMP" || true

# 3. Append (or replace) this version's { version, digest } line.
node .github/scripts/v4/generate-v4-tag-list.js \
  --zip "$ZIP" \
  --version "$VERSION" \
  --ndjson "$NDJSON"

# 4. Upsert the tag-list release and publish the merged NDJSON.
if ! gh release view template-v4-tag-list >/dev/null 2>&1; then
  gh release create template-v4-tag-list \
    --title "Template v4 Tag List" \
    --notes "NDJSON tag list for the v4 template channel."
fi
gh release upload template-v4-tag-list "$NDJSON" --clobber
