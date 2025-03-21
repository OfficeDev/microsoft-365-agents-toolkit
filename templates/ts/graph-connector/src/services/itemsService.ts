import { getAllItemsFromAPI } from "../custom/getAllItemsFromAPI";
import { Config } from "../models/Config";
import { Item } from "../models/Item";

/**
 * Gets all items from the repository.
 * @param config - The configuration object.
 * @returns An array of items.
 */
export async function getAllItems(config: Config, since?: Date): Promise<Item[]> {
  return getAllItemsFromAPI(config, since);
}
