import { Item } from "../models/Item";
import { ExternalConnectors } from "@microsoft/microsoft-graph-types";
import { getAclFromITem } from "./getAclFromItem";

export function getExternalItemFromItem(item: Item): ExternalConnectors.ExternalItem {
  return {
    id: item.id,
    properties: {
      lastModified: item.lastModified,
      "title@odata.type": "String",
      title: item.title,
      "abstract@odata.type": "String",
      abstract: item.abstract,
      assignedTo: item.assignedTo,
      issueNumber: item.issueNumber.toString(),
      state: item.state,
      author: item.author,
      url: item.url,
      owner: item.owner,
      repo: item.repo,
    },
    content: {
      value: item.content,
      type: "text",
    },
    acl: getAclFromITem(item),
  } as ExternalConnectors.ExternalItem;
}
