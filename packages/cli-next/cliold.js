#!/usr/bin/env node

/**
 * Copyright (c) Microsoft Corporation.
 * Licensed under the MIT license.
 */

process.env.TEAMSFX_CLI_BIN_NAME = "teamsapp";
require("./build/index").start();
