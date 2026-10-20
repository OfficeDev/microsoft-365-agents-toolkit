const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const test = require("node:test");

const dockerfile = fs.readFileSync(path.join(__dirname, "Dockerfile"), "utf8");

test("Chrome wrapper uses test mode without replacing the first-run flag", () => {
  assert.ok(
    dockerfile.includes(
      "sed -i 's/--no-first-run/--no-first-run --test-type/' /usr/local/bin/chrome-wrapper",
    ),
    "Chrome test launches must suppress startup warning bars while retaining --no-first-run",
  );
});

test("Chrome desktop actions launch the existing first-run-safe wrapper", () => {
  assert.ok(
    dockerfile.includes(
      "sed -i 's|^Exec=/usr/bin/google-chrome-stable|Exec=/usr/local/bin/chrome-wrapper|' /usr/share/applications/google-chrome.desktop",
    ),
    "Chrome desktop actions must not bypass chrome-wrapper",
  );
});

test("HTTP and HTTPS browser associations belong to the VS Code user", () => {
  assert.ok(
    /runuser -u vscode -- env HOME=\/home\/vscode XDG_CONFIG_HOME=\/home\/vscode\/\.config xdg-mime default google-chrome\.desktop x-scheme-handler\/http x-scheme-handler\/https/.test(
      dockerfile.replace(/\\\r?\n\s*/g, "").replace(/\s+/g, " "),
    ),
    "HTTP and HTTPS must use Chrome's desktop entry for the vscode user",
  );
});
