import { BaseDashboard } from "@microsoft/teamsfx-react";

import ChartWidget from "../widgets/ChartWidget";
import ListWidget from "../widgets/ListWidget";

export default class SampleDashboard extends BaseDashboard {
  layout() {
    return (
      <>
        <ListWidget />
        <ChartWidget />
      </>
    );
  }
}
