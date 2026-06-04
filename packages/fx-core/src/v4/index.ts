// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * Public surface of the v4 scaffolding world.
 *
 * v4 lives isolated from v3 (see scaffolding.create.proposal.md §5.1). It
 * imports no v3 symbol; v3 may call into this barrel, but nothing here is
 * tailored for v3.
 */

export * from "./distribution/templateSource";
export * from "./distribution/templateSourcePort";
export * from "./distribution/bundledFloor";
export * from "./distribution/templateConfig";
export * from "./distribution/templatePackage";
