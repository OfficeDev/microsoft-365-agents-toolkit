// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import type { UserInteraction, IProgressHandler } from "../api/qm/ui";
import type { LifecycleName, LifecycleProgress } from "./types";

/**
 * Create a LifecycleProgress adapter that delegates to the platform's
 * UserInteraction progress bar (VS Code notification, CLI spinner, etc.).
 *
 * If the consumer doesn't supply their own progress callbacks, operations
 * use this adapter to show native progress UI automatically.
 */
export function createProgressAdapter(ui: UserInteraction, title?: string): LifecycleProgress {
  let handler: IProgressHandler | undefined;

  return {
    async onStart(lifecycle: LifecycleName, totalSteps: number): Promise<void> {
      const label = title ?? `Running ${lifecycle}`;
      handler = ui.createProgressBar(label, totalSteps);
      await handler.start();
    },

    async onStepStart(_stepIndex: number, stepLabel: string): Promise<void> {
      await handler?.next(stepLabel);
    },

    onStepComplete(): void {
      // Progress bar already advanced in onStepStart via next()
    },

    async onEnd(success: boolean): Promise<void> {
      await handler?.end(success);
    },
  };
}

/**
 * A no-op LifecycleProgress implementation for silent execution
 * (e.g. non-interactive CI pipelines or tests that don't need progress).
 */
export const silentProgress: LifecycleProgress = {
  onStart() {},
  onStepStart() {},
  onStepComplete() {},
  onEnd() {},
};
