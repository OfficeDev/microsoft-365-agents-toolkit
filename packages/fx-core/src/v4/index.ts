// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * Public surface of the v4 scaffolding world.
 *
 * v4 lives isolated from v3 (see scaffolding.create.proposal.md §5.1). It
 * imports no v3 symbol; v3 may call into this barrel, but nothing here is
 * tailored for v3.
 */

export * from "./model/dataModel";
export * from "./distribution/templateSource";
export * from "./distribution/templateSourcePort";
export * from "./distribution/bundledFloor";
export * from "./distribution/templateConfig";
export * from "./distribution/templatePackage";
export * from "./distribution/packageDir";
export * from "./distribution/declarativePackage";
export * from "./distribution/createSelector";
export * from "./distribution/createQuestions";
export * from "./expression/evaluateExpression";
export * from "./renderContext/buildRenderContext";
export * from "./collectInputs/collectInputs";
export * from "./pipeline/runScaffoldPipeline";
export * from "./validation/validateTemplatePackage";
export * from "./buildTarget/resolveBuildTarget";
export * from "./buildTarget/parseSelector";
export * from "./runtime/whitelist";
export * from "./runtime/renderMustache";
export * from "./runtime/steps/mcpAuth";
export * from "./runtime/packageParse";
export * from "./runtime/inMemoryRuntime";
export * from "./runtime/realRuntime";
export * from "./runtime/scaffold";
export * from "./runtime/scaffoldFromPackageDir";
export * from "./surface/uiPromptUI";
export * from "./surface/createInputs";
export * from "./surface/createSelectorWalk";
