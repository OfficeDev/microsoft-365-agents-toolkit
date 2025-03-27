import { ItemsExtractorMaybeAsync, NextPageUrlExtractorMaybeAsync, PagedItemsService } from "../services/itemsService";
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
 * Generic service that calls the fetch API on a specified URL.
 * 
 * Adds paging and item transform capabilities.
 */
export class UrlFetchService<T> implements PagedItemsService<T> {
  url: string;
  options?: RequestInit;
  nextPageUrl?: string;
  itemsExtractor?: ItemsExtractorMaybeAsync<Response, T[]>;
  nextPageExtractor?: NextPageUrlExtractorMaybeAsync<Response>;

  constructor({ url, init, itemsExtractor, nextPageExtractor }: UrlFetchServiceParameters<T>) {
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
      throw new Error(`Failed to fetch from URL ${url}: ${response.statusText}`);
    }
    // check if response has next page
    this.nextPageUrl = await asPromise(this.nextPageExtractor?.call(response));
    return asPromise(this.itemsExtractor(response));
  }
}
