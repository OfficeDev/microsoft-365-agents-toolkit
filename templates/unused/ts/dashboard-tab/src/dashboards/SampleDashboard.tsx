import { BaseDashboard } from "./BaseDashboard";

import ChartWidget from "../widgets/ChartWidget";
import ListWidget from "../widgets/ListWidget";

export default class SampleDashboard extends BaseDashboard<any, any> {
  override layout(): JSX.Element | undefined {
    return (
      <>
        <ListWidget />
        <ChartWidget />
      </>
    );
  }
}
