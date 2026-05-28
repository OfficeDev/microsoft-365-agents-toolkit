// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { TelemetryTriggerFrom } from "../../telemetry/extTelemetryEvents";

export type SampleGalleryState = {
  loading: boolean;
  filteredSamples?: Array<SampleInfo>;
  error?: Error;
  selectedSampleId?: string;

  // keep filtering state here to recover after navigating back from detail page
  layout: "grid" | "list";
  query: string;
  filterTags: string[];
};

export type SampleDetailState = {
  loading: boolean;
  readme: string;
  error?: Error;
};

export interface SampleInfo {
  id: string;
  title: string;
  shortDescription: string;
  fullDescription: string;
  types: string[];
  tags: string[];
  time: string;
  configuration: string;
  suggested: boolean;
  downloadUrlInfo: {
    owner: string;
    repository: string;
    ref: string;
    dir: string;
  };
  thumbnailPath: string;
  gifUrl?: string;
  // -1 means TTK is lower than required.
  versionComparisonResult: -1 | 0 | 1;
  minimumToolkitVersion?: string;
  maximumToolkitVersion?: string;
}

export type SampleProps = {
  sample: SampleInfo;
  selectSample: (id: string, triggerFrom: TelemetryTriggerFrom) => void;
  createSample: (sample: SampleInfo, triggerFrom: TelemetryTriggerFrom) => void;
  viewGitHub: (sample: SampleInfo, triggerFrom: TelemetryTriggerFrom) => void;
  upgradeToolkit: (sample: SampleInfo, triggerFrom: TelemetryTriggerFrom) => void;
};

export type SampleFilterOptionType = {
  capabilities: string[];
  languages: string[];
  technologies: string[];
};

export type SampleFilterProps = {
  samples: Array<SampleInfo>;
  filterOptions: SampleFilterOptionType;
  layout: "grid" | "list";
  query: string;
  filterTags: string[];

  onLayoutChanged: (layout: "grid" | "list") => void;
  onFilterConditionChanged: (query: string, filterTags: string[]) => void;
};
