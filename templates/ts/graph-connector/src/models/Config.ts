import { InvocationContext } from "@azure/functions";
import { ExternalConnectors } from "@microsoft/microsoft-graph-types";

export interface Config {
  context: InvocationContext;
  clientId: string;
  connector: {
    accessToken: string;
    id: string;
    name: string;
    description: string;
    schema: ExternalConnectors.Schema;
    template: any;
    repos: string;
  };
}
