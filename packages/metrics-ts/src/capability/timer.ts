// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import { performance } from "perf_hooks";
import { timerData, TimerSeverity, tracePoint } from "../rawData";
import { traceId } from "../tracing";
import { appendOutput, appendOutputSync } from "../writer";

/**
 * TODO: allow cutomization
 * all in millsecond
 */
const fast = 1;
const normal = 20;

export const MSTimer = (fn: string) => {
  /* eslint-disable  @typescript-eslint/no-explicit-any */

  return (target: any, propertyKey: string, descriptor: PropertyDescriptor) => {
    const originalMethod = descriptor.value;

    /**
     * get essential properties
     */
    const data: tracePoint = {
      traceId: traceId(),
      file: fn,
      class: target.constructor.name,
      method: originalMethod.name,
      timestamp: Date.now(),
    };

    if (originalMethod.constructor.name === "AsyncFunction") {
      descriptor.value = async function (...args: any[]) {
        const start = performance.now();
        const result = await originalMethod.apply(this, args);
        const end = performance.now();
        const duration = end - start;
        let severity: TimerSeverity;
        if (duration > normal) {
          severity = TimerSeverity.Slow;
        } else if (duration > fast) {
          severity = TimerSeverity.Normal;
        } else {
          severity = TimerSeverity.Fast;
        }
        const timerData: timerData = {
          duration: duration,
          severity: severity,
        };

        data.timer = timerData;

        await appendOutput(data);
        return result;
      };
    } else {
      descriptor.value = function (...args: any[]) {
        const start = performance.now();
        const result = originalMethod.apply(this, args);
        const end = performance.now();
        const duration = end - start;
        let severity: TimerSeverity;
        if (duration > normal) {
          severity = TimerSeverity.Slow;
        } else if (duration > fast) {
          severity = TimerSeverity.Normal;
        } else {
          severity = TimerSeverity.Fast;
        }
        const timerData: timerData = {
          duration: duration,
          severity: severity,
        };

        data.timer = timerData;

        appendOutputSync(data);
        return result;
      };
    }
    return descriptor;
  };
};
