// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
export interface tracePoint {
  traceId: string;
  //pspanId: number;
  //spanId: number;
  file: string;
  class: string;
  timestamp: number;
  method: string;
  /* eslint-disable  @typescript-eslint/no-explicit-any */
  args?: any;
  timer?: timerData;
}

export enum TimerSeverity {
  Fast = "fast",
  Normal = "normal",
  Slow = "slow",
}

export interface timerData {
  duration: number;
  severity: TimerSeverity;
}
