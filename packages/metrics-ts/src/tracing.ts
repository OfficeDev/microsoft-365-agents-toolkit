// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import { env } from "process";
import { v4 as uuidv4 } from "uuid";

/**
 * TODO: make it safe
 */
export function traceId(): string {
  if (env.METAOS_TRACE_ID) {
    return env.METAOS_TRACE_ID;
  } else {
    const traceId = uuidv4();
    env.METAOS_TRACE_ID = traceId;
    return traceId;
  }
}
