const path = require("path");
const os = require("os");
const crypto = require("crypto");
const fs = require("fs");
const m = require("./build");
m.registerBuiltinTemplates();
m.registerBuiltinDrivers();
const testDir = path.join(os.tmpdir(), "atk-scaffold-check-" + Date.now());
fs.mkdirSync(testDir, { recursive: true });
const appName = "testCheck";
const ctx = m.createAtkContext({
  auth: { m365TokenProvider: {}, azureAccountProvider: {} },
  logger: {
    log: () => {},
    verbose: () => {},
    debug: (...args) => console.log("[debug]", ...args),
    info: () => {},
    warning: console.warn,
    error: console.error,
    logInFile: async () => {},
    getLogFilePath: () => "/dev/null",
  },
  telemetry: {
    sendTelemetryEvent: () => {},
    sendTelemetryErrorEvent: () => {},
    sendTelemetryException: () => {},
  },
  ui: {},
  projectPath: testDir,
  correlationId: crypto.randomUUID(),
});
m.runOperation(m.project.createProjectOp, ctx, {
  templateId: "bot/echo",
  projectName: appName,
  language: "typescript",
  destinationPath: testDir,
})
  .then((result) => {
    if (result.isErr()) {
      console.log("SCAFFOLD FAILED:", result.error.name, result.error.message);
      return;
    }
    const projectPath = result.value.projectPath;
    console.log("Scaffolded to:", projectPath);
    function listFiles(dir, prefix) {
      try {
        const entries = fs.readdirSync(dir, { withFileTypes: true });
        for (const entry of entries) {
          const fullPath = path.join(dir, entry.name);
          const relativePath = prefix ? prefix + "/" + entry.name : entry.name;
          if (entry.isDirectory()) {
            listFiles(fullPath, relativePath);
          } else {
            console.log("  " + relativePath);
          }
        }
      } catch (e) {
        console.log("Error listing", dir, e.message);
      }
    }
    console.log("\nFiles in scaffolded project:");
    listFiles(projectPath, "");
    const m365yml = path.join(projectPath, "m365agents.yml");
    const teamsappYml = path.join(projectPath, "teamsapp.yml");
    console.log("\nm365agents.yml exists:", fs.existsSync(m365yml));
    console.log("teamsapp.yml exists:", fs.existsSync(teamsappYml));
    const envDir = path.join(projectPath, "env");
    console.log("env/ dir exists:", fs.existsSync(envDir));
    if (fs.existsSync(envDir)) {
      console.log("env/ contents:", fs.readdirSync(envDir));
    }
  })
  .catch((e) => console.error("EXCEPTION:", e));
