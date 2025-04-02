// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

const fs = require("fs-extra");
const path = require("path");
const pkg = require("pkg");
const pkg_fetch = require("pkg-fetch");
const rcedit = require("rcedit");
const { version } = require("./package.json");

process.env.PKG_CACHE_PATH = path.resolve("./.pkg-cache");

const target = "node18-win-x64";
const [nodeRange, platform, arch] = target.split("-");

(async () => {
  console.log(process.env.PKG_CACHE_PATH);
  const fetched = await pkg_fetch.need({ nodeRange, platform, arch });
  console.log(fetched);
  await new Promise((resolve) => setTimeout(resolve, 5000));
  if (fetched) {
    const dirname = path.dirname(fetched);
    const filename = path.basename(fetched);
    await rcedit(fetched, {
      "version-string": {
        CompanyName: "Microsoft",
        LegalCopyright: "\xA9 Microsoft Corporation. All rights reserved.",
        ProductName: "Microsoft 365 Agents Toolkit",
        FileDescription: "Microsoft 365 Agents Toolkit Server",
        InternalFilename: "server",
        OriginalFilename: "server.exe",
      },
      "file-version": version,
      "product-version": version,
      "requested-execution-level": "asInvoker",
      // "application-manifest": "",
    });
    const built = path.join(dirname, filename.replace("fetched", "built"));
    await fs.rename(fetched, built);

    const output = path.join("lib", `server-${arch}-${version}.exe`);
    await pkg.exec(["./lib/index.js", "-t", target, "-o", output, "-c", "pkg.json", "--build"]);
    console.log("List files in lib folder:");
    fs.readdir("./lib/", (err, files) => {
      files.forEach((file) => {
        console.log("\t" + file);
      });
    });
  }
})();
