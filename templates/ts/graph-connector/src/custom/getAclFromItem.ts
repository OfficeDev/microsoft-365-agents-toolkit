import { Item } from "../models/Item";
import { ExternalConnectors } from "@microsoft/microsoft-graph-types";

// [Customization point]
// If you need additional logic to set the ACL for the item, you can add it here
// This function is used to set the ACL for each item
// in the Graph API. The ACL is used to control access to the item.
// For example, you can use a different ACL for different items, etc.
export function getAclFromITem(item: Item): ExternalConnectors.Acl[] {
  return [
    {
      accessType: "grant",
      type: "everyone",
      value: "everyone",
    },
  ];
}
