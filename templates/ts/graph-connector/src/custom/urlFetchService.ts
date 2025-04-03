import { ItemsExtractorMaybeAsync, ItemsService, NextPageUrlExtractorMaybeAsync, PagedItemsService, ProcessArgs, ResponseBasedPagingParameters } from "../services/itemsService";
import { asPromise } from "../utils";

/**
 * Options used to construct the URL fetch service.
 */
interface UrlFetchServiceParameters<T> {
  url: string,
  init?: RequestInit,
  itemsExtractor: ItemsExtractorMaybeAsync<Response, T[]>;
  nextPageExtractor?: NextPageUrlExtractorMaybeAsync<Response>;
}

/**
 * Extracts items from a JSON response.
 * 
 * @param response The Fetch API response object.
 * @returns A promise that resolves to the items
 */
export async function extractItemsFromJsonResponse<T>(response: Response): Promise<T> {
  return await response.json();
}

/**
 * Gets the next page's link from the HTTP Link response header.
 * 
 * @param response The Fetch API response object.
 * @returns The next page's URL or null if there's no next page.
 */
export function nextPageFromLinkResponseHeaderSync({response}: ResponseBasedPagingParameters<Response>): string | null {
  // Format outlined in: https://developer.mozilla.org/en-US/docs/Web/HTTP/Reference/Headers/Link#pagination_through_links
  let nextLink = response.headers.get("link").split(",").find((link) => link.includes('rel="next"'));
  let nextPageUrl = nextLink?.match(/<(.+)>/)[1];
  return nextPageUrl
}

/**
 * Generic service that calls the fetch API on a specified URL.
 * 
 * Adds paging and item transform capabilities.
 */
export class UrlFetchService<T> implements PagedItemsService<T>, ItemsService<T> {
  url: string;
  options?: RequestInit;
  nextPageUrl?: string;
  itemsExtractor: ItemsExtractorMaybeAsync<Response, T[]>;
  nextPageExtractor?: NextPageUrlExtractorMaybeAsync<Response>;

  constructor(
    {
      url,
      init,
      itemsExtractor = extractItemsFromJsonResponse,
      nextPageExtractor = nextPageFromLinkResponseHeaderSync
    }: UrlFetchServiceParameters<T>
  ) {
    this.url = url;
    this.options = init;
    this.nextPageUrl = url;
    this.itemsExtractor = itemsExtractor;
    this.nextPageExtractor = nextPageExtractor;
  }

  hasNextPage(): boolean { return !!this.nextPageUrl; }
  async getNextPageAsync(): Promise<T[]> {
    const url = this.nextPageUrl;
    if (!url) {
      // no url
      return [];
    }
    const response = await fetch(url, this.options);
    if (!response.ok) {
      throw new Error(`Failed to fetch from URL ${url}: ${response?.statusText}`);
    }
    // check if response has next page
    if (this.nextPageExtractor) {
      this.nextPageUrl = await asPromise(this.nextPageExtractor({response}));
    } else {
      this.nextPageUrl = null;
    }
    return asPromise(this.itemsExtractor(response));
  }
  async processAllAsync({ processor }: ProcessArgs<T>): Promise<void> {
    while (this.hasNextPage()) {
      var page = await this.getNextPageAsync();
      await asPromise(processor(page));
    }
  }
}
