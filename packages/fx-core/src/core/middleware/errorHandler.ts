// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
"use strict";

import { HookContext, Middleware, NextFunction } from "@feathersjs/hooks";
import { err, Inputs, SystemError, UserError } from "@microsoft/teamsfx-api";
import { setLocale } from "../../common/globalVars";
import { settingsUtil } from "../../component/utils/settingsUtil";
import { assembleError, FilePermissionError } from "../../error/common";

/**
 * in case there're some uncatched exceptions, this middleware will act as a guard
 * to catch exceptions and return specific error.
 */
export const ErrorHandlerMW: Middleware = async (ctx: HookContext, next: NextFunction) => {
  const inputs = ctx.arguments[ctx.arguments.length - 1] as Inputs;
  if (inputs?.locale) setLocale(inputs.locale);
  if (inputs?.projectPath) {
    try {
      await settingsUtil.readSettings(inputs.projectPath, false);
    } catch {
      // Ignore trackingId preload failure and continue normal flow.
    }
  }
  try {
    await next();
  } catch (e) {
    const fxError = assembleError(e);
    ctx.result = err(fxError);
  }
  if (ctx.result?.isErr()) {
    let fxError = ctx.result.error;
    fxError = convertError(fxError);
    ctx.result = err(fxError);
  }
};

const Reg1 =
  /The client '.+' with object id '.+' does not have authorization to perform action '.+' over scope '.+' or the scope is invalid. If access was recently granted, please refresh your credentials\./;
const Reg2 = /"resourceGroupName" with value ".+" should satisfy the constraint "Pattern"/;
const Reg3 = /Resource '.+' was disallowed by policy./;
const Reg4 =
  /The subscription '.+' is disabled and therefore marked as read only. You cannot perform any write actions on this subscription until it is re-enabled\./;
const Reg5 =
  /The current subscription type is not permitted to perform operations on any provider namespace. Please use a different subscription\./;
const Reg6 =
  /The provided location '.+' is not available for resource group\. List of available regions is '.+'\./;
const Reg7 = /The subscription '.+' could not be found\./;
const Reg8 =
  /Invalid resource group location '.+'. The Resource group already exists in location '.+'\./;
const Reg9 = /The access token is from the wrong issuer '.+'\./;
const Reg10 = /Entry not found in cache\./;
const Reg11 = /request to .+ failed, reason: .+/;
const Reg12 = /.+no space left on device.+/;
// const Reg12 = /ENOENT: no such file or directory/;
// const Reg13 = /EBUSY: resource busy or locked/;
// const Reg14 = /Lock is not .+ by you/;
const Reg15 = /EPERM: operation not permitted/;

const Regs = [Reg1, Reg2, Reg3, Reg4, Reg5, Reg6, Reg7, Reg8, Reg9, Reg10, Reg11, Reg12];

export function convertError(err: Error): UserError | SystemError | Error {
  const msg = err.message;
  if (!msg) return err;
  if (Reg15.test(msg) === true) {
    const userError = new FilePermissionError(err, (err as any).source);
    return userError;
  }
  for (const reg of Regs) {
    if (reg.test(msg) === true) {
      const userError = new UserError(
        (err as any).source,
        err.name,
        err.message,
        (err as any).displayMessage
      );
      userError.stack = err.stack;
      return userError;
    }
  }
  return err;
}
