// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

// applicationinsights is loaded lazily to avoid pulling its module graph
// (and its transitive dependencies) at CLI startup.  This shaves ~100ms
// off cold-start because the package is an esbuild external.
type AppInsightsModule = typeof import("applicationinsights");

/**
 * Thin wrapper around Application Insights TelemetryClient.
 * All methods are no-ops until `init()` is called with a valid key.
 */
export class AppInsightsTransport {
  private client: import("applicationinsights").TelemetryClient | undefined;

  /**
   * Initialise the App Insights client.  Must be called once before any
   * tracking.  If called with an empty key the transport stays inert.
   */
  init(instrumentationKey: string, commonProperties: Record<string, string>): void {
    if (!instrumentationKey) return;

    const appInsights: AppInsightsModule = require("applicationinsights");

    if (appInsights.defaultClient) {
      this.client = new appInsights.TelemetryClient(instrumentationKey);
      this.client.channel.setUseDiskRetryCaching(true);
    } else {
      appInsights
        .setup(instrumentationKey)
        .setAutoCollectRequests(false)
        .setAutoCollectPerformance(false)
        .setAutoCollectExceptions(false)
        .setAutoCollectDependencies(false)
        .setAutoDependencyCorrelation(false)
        .setAutoCollectConsole(false)
        .setUseDiskRetryCaching(true)
        .start();
      this.client = appInsights.defaultClient;
    }

    this.client.commonProperties = commonProperties;
  }

  trackEvent(
    name: string,
    properties?: Record<string, string>,
    measurements?: Record<string, number>
  ): void {
    this.client?.trackEvent({ name, properties, measurements });
  }

  trackException(
    error: Error,
    properties?: Record<string, string>,
    measurements?: Record<string, number>
  ): void {
    this.client?.trackException({ exception: error, properties, measurements });
  }

  flush(): Promise<void> {
    if (!this.client) return Promise.resolve();
    return new Promise<void>((resolve) => {
      this.client!.flush({ callback: () => resolve() });
    });
  }
}
