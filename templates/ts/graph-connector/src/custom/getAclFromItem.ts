import { Item } from "../models/Item";
import { ExternalConnectors } from "@microsoft/microsoft-graph-types";

export function getAclFromITem(item: Item): ExternalConnectors.Acl[] {
  return [
    {
      accessType: "grant",
      type: "everyone",
      value: "everyone",
    },
  ];
}
