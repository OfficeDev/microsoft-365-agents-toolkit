// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { ImportInfo } from "./importInfo";
import * as rules from "./mappings.json";
import { FunctionReplacement, Replacement, ImportInfoInReplacement } from "./replacement";

/**
 * generate need replaced interface map
 * @returns replacement rules lis
 */
export function buildReplacement(name: "interfaceMappings" | "enumMappings"): Replacement[] {
  return rules[name].map((item) => defaultMapping(item));
}

/**
 * generate need replaced function map
 * @returns function replacement rules list
 */
export function buildFunctionReplacement(): FunctionReplacement[] {
  return rules["functionMappings"].map((item) => functionMapping(item));
}

/** Build a mapping for replacement
 *  @param importInfo Import information
 *  @param replacements Origin replacement rules
 *  @returns directlyImport2TargetReplacements means sourceTokens contains only one
 *           namespaceImport2TargetReplacements means sourceTokens contains multiple elements
 *
 *  i.e. Import: import { A, B as aliasB } from "xxx"
 *       Mapping: A => X.A; B.I => Y.I
 *       Return:
 *        {
 *          directlyImport2TargetReplacements: {
 *            A: {
 *              sourceTokens: ["A"]
 *              targetTokens: ["X", "A"]
 *            }
 *          },
 *          namespaceImport2TargetReplacements: {
 *            aliasB.I : {
 *              sourceTokens: ["aliasB", "I"]
 *              targetTokens: ["Y", "I"]
 *            }
 *          }
 *        }
 *  i.e. Import: import * as msft from "xxx"
 *       Mapping: A => X.A; B.I => Y.I
 *       Return:
 *        {
 *          namespaceImport2TargetReplacements: {
 *            msft.A : {
 *             sourceTokens: ["msft", "A"]
 *             targetTokens: ["msft", "X", "A"]
 *            }
 *            msft.B.I : {
 *             sourceTokens: ["msft", "B", "I"]
 *             targetTokens: ["msft", "Y", "I"]
 *            }
 *          }
 *        }
 */
export function buildTargetReplacement<T extends Replacement>(
  importInfo: ImportInfo,
  replacements: T[]
): [
  directlyImport2TargetReplacements: Map<string, ImportInfoInReplacement & T>,
  namespaceImport2TargetReplacements: Map<string, ImportInfoInReplacement & T>
] {
  const directlyImport2TargetReplacements: Map<string, ImportInfoInReplacement & T> = new Map();
  const namespaceImport2TargetReplacements: Map<string, ImportInfoInReplacement & T> = new Map();

  for (const replacement of replacements) {
    const importSingleModules = importInfo.importSingleExportInfo.filter(
      (info) => info.source === replacement.sourceTokens[0]
    );
    for (const importSingleModule of importSingleModules) {
      const alias = importSingleModule?.alias;
      let sourceTokens: string[] = replacement.sourceTokens;
      if (alias) {
        sourceTokens = replacement.sourceTokens.map((token, index) => {
          return index === 0 ? alias : token;
        });
      }
      const sourceRule = Object.assign({}, replacement);
      const targetRule = Object.assign(sourceRule, {
        sourceTokens: sourceTokens,
        importSingleExportInfo: importSingleModule,
      });

      if (sourceTokens.length == 1) {
        directlyImport2TargetReplacements.set(sourceTokens[0], targetRule);
      } else {
        namespaceImport2TargetReplacements.set(sourceTokens.join("."), targetRule);
      }
    }

    importInfo.importEntireModuleInfo.forEach((importEntireModuleInfo) => {
      const sourceRule = Object.assign({}, replacement);
      const targetRule = Object.assign(sourceRule, {
        sourceTokens: [importEntireModuleInfo.alias, ...replacement.sourceTokens],
        targetTokens: [importEntireModuleInfo.alias, ...replacement.targetTokens],
      });
      namespaceImport2TargetReplacements.set(targetRule.sourceTokens.join("."), targetRule);
    });
  }
  return [directlyImport2TargetReplacements, namespaceImport2TargetReplacements];
}

function functionMapping(item: {
  source: string;
  target: string;
  callbackToPromise?: boolean;
  callbackPosition?: number;
}): FunctionReplacement {
  const baseReplacement = defaultMapping(item);
  return Object.assign(baseReplacement, {
    callbackToPromise: item?.callbackToPromise ?? false,
    callbackPosition: item?.callbackPosition === undefined ? -1 : item?.callbackPosition,
  });
}

function defaultMapping(item: { source: string; target: string }): Replacement {
  const sourceTokenList = item.source.trim().split(".");
  const targetTokenList = item.target.trim().split(".");
  if (sourceTokenList.length < 2 || targetTokenList.length < 2) {
    throw new Error(`Error mapping item: ${JSON.stringify(item)}`);
  }
  return {
    sourceTokens: sourceTokenList.slice(1),
    targetTokens: targetTokenList.slice(1),
  };
}
