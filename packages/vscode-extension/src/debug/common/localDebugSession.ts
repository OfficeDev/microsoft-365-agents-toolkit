// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import * as uuid from "uuid";
import { DebugNoSessionId } from "../common/debugConstants";
import { allRunningDebugSessions } from "./globalVariables";
import VsCodeLogInstance from "../../commonlib/log";

export class LocalDebugSession {
  static createSession() {
    const session = new LocalDebugSession(uuid.v4());
    return session;
  }
  static createInvalidSession() {
    return new LocalDebugSession();
  }

  readonly id: string;
  // Save the time when the event it sent for calculating time gaps.
  readonly eventTimes: { [eventName: string]: number | undefined } = {};
  readonly properties: { [key: string]: string } = {};
  readonly errorProps: string[] = [];
  readonly failedServices: { name: string; exitCode: number | undefined }[] = [];

  private constructor(id: string = DebugNoSessionId) {
    this.id = id;
  }
}

// Helper functions for local debug correlation-id, only used for telemetry
// Use a 2-element tuple to handle concurrent F5
const localDebugCorrelationIds: [LocalDebugSession, LocalDebugSession] = [
  LocalDebugSession.createInvalidSession(),
  LocalDebugSession.createInvalidSession(),
];
let current = 0;
export function startLocalDebugSession(): string {
  current = (current + 1) % 2;
  localDebugCorrelationIds[current] = LocalDebugSession.createSession();
  return getLocalDebugSessionId();
}

export function endLocalDebugSession() {
  localDebugCorrelationIds[current] = LocalDebugSession.createInvalidSession();
  current = (current + 1) % 2;
}

export function getLocalDebugSession(): LocalDebugSession {
  return localDebugCorrelationIds[current];
}

export function getLocalDebugSessionId(): string {
  return localDebugCorrelationIds[current].id;
}

export async function checkAndSkipDebugging(): Promise<boolean> {
  // skip debugging if there is already a debug session
  if (allRunningDebugSessions.size > 0) {
    VsCodeLogInstance.warning("Skip debugging because there is already a debug session.");
    endLocalDebugSession();
    return Promise.resolve(true);
  }
  return Promise.resolve(false);
}
