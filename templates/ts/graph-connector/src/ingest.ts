import { ExternalConnectors } from "@microsoft/microsoft-graph-types";
import { getClient } from "./graphClient";
import { Config } from "./models/Config";
import { getAllItems } from "./services/itemsService";
import { Item } from "./models/Item";
import { getExternalItemFromItem } from "./custom/getExternalItemFromItem";

const client = getClient();

/**
 * Transforms the content into a format that can be ingested by the Graph API.
 * @param items - The items to transform.
 * @returns An array of objects that can be ingested by the Graph API.
 */
function transformContent(items: Item[]): ExternalConnectors.ExternalItem[] {
  return items.map((item) => {
    return getExternalItemFromItem(item);
  });
}

/**
 * Loads the content into the Graph API.
 * @param config - The configuration object.
 * @param doc - The document to load.
 * @returns A promise that resolves when the content has been loaded.
 */
async function loadContent(config: Config, item: ExternalConnectors.ExternalItem): Promise<void> {
  const itemId = item.id;

  // Remove the ID from the item to avoid conflicts
  delete item.id;

  try {
    const url = `/external/connections/${config.connector.id}/items/${itemId}`;

    config.context.log(`PUT ${url}`);
    config.context.log(JSON.stringify(item, null, 4));

    await client.api(url).header("content-type", "application/json").put(item);
  } catch (e) {
    config.context.error(`Failed to load ${itemId}: ${e.message}`);
    if (e.body) {
      config.context.error(`${JSON.parse(e.body, null)?.innerError?.message ?? ""}`);
    }
    return;
  }
}

/**
 * Ensures that the content is ingested into the Graph API.
 * @param config - The configuration object.
 */
export async function ingestContent(config: Config, since?: Date): Promise<void> {
  const files = await getAllItems(config, since);
  for (const doc of transformContent(files)) {
    await loadContent(config, doc);
  }
}
