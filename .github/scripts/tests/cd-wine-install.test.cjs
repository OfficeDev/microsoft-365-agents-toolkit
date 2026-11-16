const assert = require("node:assert/strict");
const { test } = require("node:test");
const fs = require("node:fs");
const path = require("node:path");

const workflow = fs.readFileSync(
  path.resolve(__dirname, "..", "..", "workflows", "cd.yml"),
  "utf8"
).replaceAll("\r\n", "\n");
const install = workflow.split("      - name: Install wine64\n")[1]
  .split("\n      - name:")[0];

test("CD pins the Ubuntu release used by the WineHQ repository", () => {
  assert.match(workflow, /^  cd:\n    runs-on: ubuntu-24\.04$/m);
});

test("WineHQ packages target the pinned runner's noble distribution", () => {
  assert.match(install, /https:\/\/dl\.winehq\.org\/wine-builds\/ubuntu\/dists\/noble\/winehq-noble\.sources/);
  assert.doesNotMatch(install, /jammy/);
});

test("the armored WineHQ key and Signed-By both use the supported asc path", () => {
  assert.match(install, /sudo wget -O \/etc\/apt\/keyrings\/winehq-archive\.asc https:\/\/dl\.winehq\.org\/wine-builds\/winehq\.key/);
  assert.ok(install.includes(
    "sudo sed -i 's|^Signed-By:.*|Signed-By: /etc/apt/keyrings/winehq-archive.asc|' /etc/apt/sources.list.d/winehq-noble.sources"
  ));
});

test("Wine installation keeps signature verification and checks both entry points", () => {
  assert.match(install, /sudo dpkg --add-architecture i386/);
  assert.match(install, /sudo apt-get update/);
  assert.match(install, /sudo apt-get install --install-recommends winehq-stable -y/);
  assert.match(install, /sudo ln -s \/usr\/bin\/wine \/usr\/bin\/wine64/);
  assert.match(install, /wine --version/);
  assert.match(install, /wine64 --version/);
  assert.doesNotMatch(install, /trusted=yes|allow-unauthenticated|AllowInsecureRepositories|AllowDowngradeToInsecureRepositories/);
});
