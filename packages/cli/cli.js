#!/usr/bin/env node

// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

const chalk = require("chalk");

process.on("uncaughtException", (err) => {
  if (err.message.includes("async_hooks")) {
    console.error(
      chalk.redBright(
        "Microsoft 365 Agents Toolkit CLI requires to use node version higher than 14.x, please update your node version."
      )
    );
  } else {
    console.error(err);
  }
  process.exit(1);
});
process.env.TEAMSFX_CLI_BIN_NAME = "atk";
const cli = require("./lib");
cli.start();
