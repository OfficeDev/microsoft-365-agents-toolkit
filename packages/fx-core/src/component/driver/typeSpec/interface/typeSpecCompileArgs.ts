// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

export interface TypeSpecCompileArgs {
  path: string; // The path to the main.tsp file.
  manifestPath: string; // The path to the teams app manifest file.
  outputDir: string; // The directory to output the compiled files.
  typeSpecConfigPath: string; // The path to the TypeSpec config file.
}
