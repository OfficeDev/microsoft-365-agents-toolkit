const path = require("path");
const semver = require("semver");
const fs = require("fs");

const templatePath = path.join(__dirname, "../../templates");
const templateVersion = require(path.join(
  templatePath,
  "package.json"
)).version;

const fxCorePath = path.join(__dirname, "../../packages/fx-core");
const templateConfigFile = path.join(
  fxCorePath,
  "./src/common/templates-config.json"
);
const templateConfigs = JSON.parse(fs.readFileSync(templateConfigFile));
const command = process.argv[2];

if (command === 'syncVersion') {
  syncTemplateVersion(templateVersion, templateConfigs);
} else if (command === 'updateUseLocalFlag') {
  updateUseLocalFlag(templateVersion, templateConfigs)
} else {
  console.warn(
    `================== invalid command '${command}' for fxcore-sync-up-version.js ==================`
  )
}

fs.writeFileSync(templateConfigFile, JSON.stringify(templateConfigs, null, 4));

function syncTemplateVersion(templateVersion, templateConfigs) {
  console.log(
    `================== template version: ${templateVersion} ==================`
  );

  console.log(
    `================== template version in fx-core configurate as ${templateConfigs.version} ==================`
  );
  templateConfigs.localVersion = templateVersion;
  if (!semver.prerelease(templateVersion)) {
    if (!semver.intersects(templateConfigs.version, templateVersion)) {
      const parsedTemplateVersion = semver.parse(templateVersion);
      const templateVersionRange = `~${parsedTemplateVersion.major}.${parsedTemplateVersion.minor}`;
      console.log(
        `================== template config version is not match with template latest release version, need bump up config version ${templateVersionRange} ==================`
      );
      templateConfigs.version = templateVersionRange;
    }
  } else if (templateVersion.includes("rc")) {
    console.log("sync up template in fx-core as 0.0.0-rc");
    templateConfigs.version = "0.0.0-rc";
  }
}

function updateUseLocalFlag(templateVersion, templateConfigs) {
  if (!semver.prerelease(templateVersion) || templateVersion.includes("rc") || process.env.VS_RELEASE === "true") {
    templateConfigs.useLocalTemplate = false;
  } else {
    templateConfigs.useLocalTemplate = true;
  }
  console.log(`================== configure useLocalFlag: ${templateConfigs.useLocalTemplate} ==================`)
}
