// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { Disposable, Uri } from "vscode";
import { uriEventHandler } from "../uriHandler";

export async function getExchangeCode(): Promise<string> {
  let uriEventListener: Disposable;
  return new Promise((resolve: (value: string) => void, reject) => {
    uriEventListener = uriEventHandler.event((uri: Uri) => {
      try {
        /* eslint-disable @typescript-eslint/no-unsafe-assignment, @typescript-eslint/no-unsafe-member-access */
        const query = parseQuery(uri);
        const code = query.code;
        /* eslint-enable @typescript-eslint/no-unsafe-assignment, @typescript-eslint/no-unsafe-member-access */

        resolve(code);
      } catch (err) {
        reject(err);
      }
    });
  })
    .then((result) => {
      uriEventListener.dispose();
      return result;
    })
    .catch((err) => {
      uriEventListener.dispose();
      throw err;
    });
}

/* eslint-disable @typescript-eslint/no-unsafe-return, @typescript-eslint/no-explicit-any, @typescript-eslint/no-unsafe-member-access */
function parseQuery(uri: Uri): any {
  return uri.query.split("&").reduce((prev: any, current) => {
    const queryString: string[] = current.split("=");
    prev[queryString[0]] = queryString[1];
    return prev;
  }, {});
}
/* eslint-enable @typescript-eslint/no-unsafe-return, @typescript-eslint/no-explicit-any, @typescript-eslint/no-unsafe-member-access */
