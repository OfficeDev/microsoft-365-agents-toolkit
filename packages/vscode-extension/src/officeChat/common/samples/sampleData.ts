// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

export class SampleData {
  docLink: string;
  codeSample: string;
  description: string;
  name: string;
  definition: string;
  usage: string;

  constructor(
    name: string,
    docLink: string,
    codeSample: string,
    description: string,
    definition: string,
    usage: string
  ) {
    this.docLink = docLink;
    this.codeSample = codeSample;
    this.description = description;
    this.name = name;
    this.definition = definition;
    this.usage = usage;
  }
}
