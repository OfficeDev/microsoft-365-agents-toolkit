// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { CoreCallbackEvent, FxError } from "@microsoft/teamsfx-api";

export type CoreCallbackFunc = (name: string, err?: FxError, data?: any) => void | Promise<void>;

export class CallbackRegistry {
  private static registry: Map<CoreCallbackEvent, CoreCallbackFunc[]> = new Map();

  public static has(event: CoreCallbackEvent): boolean {
    return this.registry.has(event);
  }

  public static set(event: CoreCallbackEvent, func: CoreCallbackFunc): void {
    if (!this.registry.has(event)) {
      this.registry.set(event, []);
    }
    const funcs = this.registry.get(event) as CoreCallbackFunc[];
    funcs.push(func);
    this.registry.set(event, funcs);
  }

  public static get(event: CoreCallbackEvent): CoreCallbackFunc[] {
    if (this.registry.has(event)) {
      return this.registry.get(event)!;
    } else {
      return [];
    }
  }

  public static refresh(): void {
    this.registry = new Map();
  }
}
