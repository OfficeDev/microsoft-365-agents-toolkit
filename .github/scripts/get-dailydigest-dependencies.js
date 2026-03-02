const axios = require("axios");
const glob = require("glob");
const fs = require("fs");
const path = require("path");

const repoRoot = __dirname + "/../..";

const codeOwnerMap = new Map([
  ["copilot-plugin-from-scratch", "Ning.Tang@microsoft.com"],
  ["copilot-gpt-basic", "Ning.Tang@microsoft.com"],
  ["copilot-gpt-from-scratch-plugin", "Ning.Tang@microsoft.com"],
  ["dashboard-tab", "Ning.Tang@microsoft.com"],
  ["non-sso-tab", "zhijie.huang@microsoft.com"],
  ["non-sso-tab-ssr", "yiminjin@microsoft.com"],
  ["sso-tab", "zhijie.huang@microsoft.com"],
  ["sso-tab-ssr", "yiminjin@microsoft.com"],
  ["default-bot", "yukundong@microsoft.com"],
  ["link-unfurling", "yukundong@microsoft.com"],
  ["message-extension-action", "yukundong@microsoft.com"],
  ["message-extension-search", "yukundong@microsoft.com"],
  ["message-extension-copilot", "yukundong@microsoft.com"],
  ["non-sso-tab-default-bot/tab", "yuqzho@microsoft.com"],
  ["non-sso-tab-default-bot/bot", "yuqzho@microsoft.com"],
  ["default-bot-message-extension", "yuqzho@microsoft.com"],
  ["message-extension", "yuqzho@microsoft.com"],
  ["office-addin", "huajiezhang@microsoft.com"],
  ["copilot-plugin-existing-api", "yuqzho@microsoft.com"],
  ["copilot-plugin-existing-api-api-key", "yuqzho@microsoft.com"],
  ["copilot-plugin-from-scratch-api-key", "Ning.Tang@microsoft.com"],
  ["spfx-tab", "yuqzho@microsoft.com"],
  ["spfx-tab-import", "yuqzho@microsoft.com"],
  ["sso-tab-naa", "bowen.song@microsoft.com"],
  ["command-and-response", "qidon@microsoft.com"],
  ["notification-http-timer-trigger", "qidon@microsoft.com"],
  ["notification-http-trigger", "qidon@microsoft.com"],
  ["notification-restify", "qidon@microsoft.com"],
  ["notification-timer-trigger", "qidon@microsoft.com"],
  ["notification-webapi", "qidon@microsoft.com"],
  ["workflow", "qidon@microsoft.com"],
  ["m365-message-extension", "kuojianlu@microsoft.com"],
  ["ai-bot", "kuojianlu@microsoft.com"],
  ["ai-assistant-bot", "kuojianlu@microsoft.com"],
  ["notification-http-timer-trigger-isolated", "tianyuan@microsoft.com"],
  ["notification-http-trigger-isolated", "tianyuan@microsoft.com"],
  ["notification-timer-trigger-isolated", "tianyuan@microsoft.com"],
  ["custom-copilot-assistant-assistants-api", "kuojianlu@microsoft.com"],
  ["custom-copilot-assistant-new", "kuojianlu@microsoft.com"],
  ["custom-copilot-basic", "kuojianlu@microsoft.com"],
  ["custom-copilot-rag-custom-api", "kuojianlu@microsoft.com"],
  ["api-plugin-from-scratch", "Ning.Tang@microsoft.com"],
  ["api-message-extension-sso", "Ning.Tang@microsoft.com"],
  ["custom-copilot-rag-microsoft365", "tianyuan@microsoft.com"],
  ["custom-copilot-rag-customize", "tianyuan@microsoft.com"],
  ["custom-copilot-rag-azure-ai-search", "tianyuan@microsoft.com"],
]);

async function getTemplatesDependencies() {
  var dependenciesMap = new Map();
  const templatePkgJsonPath = `${repoRoot}/templates/**/package.json.tpl`;
  const packageJsonFiles = await glob.glob(templatePkgJsonPath, {
    ignore: "node_modules/**",
  });

  packageJsonFiles.forEach((packageJsonFile) => {
    const packageJson = JSON.parse(fs.readFileSync(packageJsonFile, "utf8"));
    let templateDir = path.relative(
      `${repoRoot}/templates`,
      path.dirname(packageJsonFile)
    );
    let codeOwners = "";
    for (const [key, value] of codeOwnerMap) {
      if (templateDir.includes("non-sso-tab-default-bot")) {
        codeOwners = "yuqzho@microsoft.com";
        continue;
      }
      if (key === path.basename(templateDir)) {
        codeOwners = value;
      }
    }
    let dependencies = packageJson["dependencies"] ?? {};
    Object.assign(dependencies, packageJson["devDependencies"]);
    for (dependency in dependencies) {
      if (
        dependenciesMap.has(dependency) &&
        dependenciesMap.get(dependency).has(codeOwners)
      ) {
        dependenciesMap.get(dependency).get(codeOwners).push({
          templateDir,
          version: dependencies[dependency],
        });
      } else if (dependenciesMap.has(dependency)) {
        dependenciesMap
          .get(dependency)
          .set(codeOwners, [
            { templateDir, version: dependencies[dependency] },
          ]);
      } else {
        const codeOwnerTemplateMap = new Map([
          [codeOwners, [{ templateDir, version: dependencies[dependency] }]],
        ]);
        dependenciesMap.set(dependency, codeOwnerTemplateMap);
      }
    }
  });

  return dependenciesMap;
}

async function getCsharpTemplateDependencies() {
  var dependenciesMap = new Map();
  const templateCsprojPath = `${repoRoot}/templates/**/*.csproj.tpl`;
  const csprojFiles = await glob.glob(templateCsprojPath);
  csprojFiles.forEach((csprojFile) => {
    const csproj = fs.readFileSync(csprojFile, "utf8");
    let templateDir = path.relative(
      `${repoRoot}/templates`,
      path.dirname(csprojFile)
    );
    let codeOwners = "";
    for (const [key, value] of codeOwnerMap) {
      if (
        path.basename(templateDir) === "copilot-plugin-existing-api" ||
        path.basename(templateDir) === "copilot-plugin-existing-api-api-key" ||
        path.basename(templateDir) === "copilot-plugin-from-oai-plugin"
      ) {
        codeOwners = "zhiyu.you@microsoft.com";
        continue;
      }
      if (key === path.basename(templateDir)) {
        codeOwners = value;
      }
    }

    const dependencies = csproj
      .split("\n")
      .filter((line) => line.includes("<PackageReference"))
      .map((line) => {
        const name = line.match(/Include="(.*)" Version/)[1];
        const version = line.match(/Version="(.*)"/)[1];
        return { name, version };
      });

    for (dependency in dependencies) {
      const dependencyName = dependencies[dependency].name;
      const dependencyVersion = dependencies[dependency].version;
      if (
        dependenciesMap.has(dependencyName) &&
        dependenciesMap.get(dependencyName).has(codeOwners)
      ) {
        dependenciesMap.get(dependencyName).get(codeOwners).push({
          templateDir,
          version: dependencyVersion,
        });
      } else if (dependenciesMap.has(dependencyName)) {
        dependenciesMap
          .get(dependencyName)
          .set(codeOwners, [{ templateDir, version: dependencyVersion }]);
      } else {
        const codeOwnerTemplateMap = new Map([
          [codeOwners, [{ templateDir, version: dependencyVersion }]],
        ]);
        dependenciesMap.set(dependencyName, codeOwnerTemplateMap);
      }
    }
  });

  return dependenciesMap;
}

function generateAdaptiveCardColumnSets(arr) {
  if (arr.length === 0) {
    return [];
  }
  let columnSets = [
    {
      type: "ColumnSet",
      columns: [
        {
          type: "Column",
          width: 22,
          items: [
            {
              type: "TextBlock",
              text: "Package",
              wrap: true,
              weight: "Bolder",
            },
          ],
          verticalContentAlignment: "Center",
        },
        {
          type: "Column",
          width: 38,
          items: [
            {
              type: "TextBlock",
              text: "Templates",
              wrap: true,
              weight: "Bolder",
            },
          ],
          verticalContentAlignment: "Center",
        },
        {
          type: "Column",
          width: 17,
          items: [
            {
              type: "TextBlock",
              text: "Version",
              wrap: true,
              weight: "Bolder",
            },
          ],
          verticalContentAlignment: "Center",
        },
        {
          type: "Column",
          width: 23,
          items: [
            {
              type: "TextBlock",
              text: "Owners",
              wrap: true,
              weight: "Bolder",
            },
          ],
          verticalContentAlignment: "Center",
        },
      ],
      separator: true,
    },
  ];
  for (package of arr) {
    let ownerColumnSets = [];
    package.ownerMap.forEach(function (templatesInfo, owner) {
      ownerColumnSets.push({
        type: "ColumnSet",
        separator: true,
        columns: [
          {
            type: "Column",
            width: 56,
            items: templatesInfo.map((templateInfo) => {
              return {
                type: "ColumnSet",
                columns: [
                  {
                    type: "Column",
                    width: 40,
                    items: [
                      {
                        type: "TextBlock",
                        text: templateInfo.templateDir,
                        wrap: true,
                        size: "Small",
                      },
                    ],
                  },
                  {
                    type: "Column",
                    width: 16,
                    items: [
                      {
                        type: "TextBlock",
                        text: templateInfo.version,
                        wrap: true,
                        size: "Small",
                      },
                    ],
                  },
                ],
              };
            }),
          },
          {
            type: "Column",
            width: 24,
            items: [
              {
                type: "TextBlock",
                text: owner,
                wrap: true,
                size: "Small",
              },
            ],
          },
        ],
      });
    });

    columnSets.push({
      type: "ColumnSet",
      columns: [
        {
          type: "Column",
          width: 20,
          items: [
            {
              type: "TextBlock",
              text:
                `[${package.name}](${
                  package.ownerMap
                    .values()
                    .next()
                    .value[0].templateDir.includes("csharp")
                    ? `https://www.nuget.org/packages/${package.name}`
                    : `https://www.npmjs.com/package/${package.name}`
                })` +
                "\n\r" +
                `LTS-${package.version}`,
              wrap: true,
              size: "Small",
            },
          ],
        },
        {
          type: "Column",
          width: 80,
          items: ownerColumnSets,
        },
      ],
      separator: true,
    });
  }

  return columnSets;
}

async function main() {
  const dependenciesMap = await getTemplatesDependencies();
  let arr = [];
  for (const entry of dependenciesMap.entries()) {
    await axios
      .get(`https://registry.npmjs.org/${entry[0]}`)
      .then((response) => {
        const ltsVersion = response.data["dist-tags"].latest;
        const ltsVersionTime = response.data.time[ltsVersion];
        const timeDiff = (new Date() - new Date(ltsVersionTime)) / 1000;
        if (timeDiff <= 86400) {
          arr.push({
            name: entry[0],
            version: ltsVersion,
            ownerMap: entry[1],
          });
        }
      });
  }

  const csharpDependencyMap = await getCsharpTemplateDependencies();
  for (const entry of csharpDependencyMap.entries()) {
    await axios
      .get(
        `https://api.nuget.org/v3/registration5-gz-semver2/${entry[0].toLowerCase()}/index.json`
      )
      .then((response) => {
        const ltsVersion = response.data["items"].at(-1).upper;
        const ltsVersionTime = response.data.commitTimeStamp;
        const timeDiff = (new Date() - new Date(ltsVersionTime)) / 1000;
        if (timeDiff <= 86400) {
          arr.push({
            name: entry[0],
            version: ltsVersion,
            ownerMap: entry[1],
          });
        }
      });
  }

  const table = generateAdaptiveCardColumnSets(arr);
  const tableString = JSON.stringify(table);
  return JSON.stringify(tableString);
}

main().then((result) => {
  console.log(result);
});
