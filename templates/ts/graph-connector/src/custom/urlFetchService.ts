import { ItemsExtractor, ItemsService, ItemsTransformer, NextPageLinkResult, NextPageUrlExtractor, PagedItemsService, PagingParameters, ProcessArgs, ResponseBasedPagingParameters, TotalItemsExtractor } from "../services/itemsService";

/**
 * Options used to construct the URL fetch service.
 */
interface UrlFetchServiceParameters<T1, T2> {
  url: string,
  init?: RequestInit,
  // Extracts items from the response
  itemsExtractor?: ItemsExtractor<Response, T1>;
  // Performs any needed transformations on the extracted items
  itemsTransformer?: ItemsTransformer<T1, T2>;
  // Extracts the total number of items from the response
  totalItemsExtractor?: TotalItemsExtractor<T1>;
  nextPageExtractor?: NextPageUrlExtractor<Response>;
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
 * Calculates the next page's link using tracked parameters.
 */
export class NextPageForOffsetPagination implements NextPageUrlExtractor<Response> {
  currentPage: number = 1;
  pageSize: number;
  pageParam: string;
  totalItems?: number;
  url: URL;
  firstPageUrl: string;

  constructor(url: string | URL, pageSize: number, pageParam: string = 'page', pageSizeParam: string = 'pageSize') {
    this.url = new URL(url);
    // set the page size.
    this.pageSize = pageSize;
    this.pageParam = pageParam;
    this.url.searchParams.set(pageSizeParam, this.pageSize.toString());
    this.firstPageUrl = this.url.toString();
  }
  firstPageLink(): NextPageLinkResult {
    return this.firstPageUrl;
  }
  nextPageLink({totalItems}: PagingParameters<Response>): NextPageLinkResult {
    if (!this.totalItems && totalItems) {
      this.totalItems = totalItems;
    }
    if (this.totalItems && (this.currentPage * this.pageSize) > this.totalItems) {
      // no next page
      return null;
    }
    const page = this.currentPage + 1;
    // set the page in the url
    this.url.searchParams.set(this.pageParam, page.toString());
    this.currentPage = page;
    return this.url.toString();
  }
}

/**
 * Gets the next page's link from the HTTP Link response header.
 */
export class NextPageFromLinkResponseHeader implements NextPageUrlExtractor<Response> {
  firstPageUrl: string;

  constructor(url: string) {
    this.firstPageUrl = url;
  }
  firstPageLink(): NextPageLinkResult {
    return this.firstPageUrl;
  }
  /**
   * Gets the next page's link from the HTTP Link response header.
   * 
   * @param response The Fetch API response object.
   * @returns The next page's URL or null if there's no next page.
   */
  nextPageLink({ response }: ResponseBasedPagingParameters<Response>): NextPageLinkResult {
    // Format outlined in: https://developer.mozilla.org/en-US/docs/Web/HTTP/Reference/Headers/Link#pagination_through_links
    let nextLink = response.headers.get("link").split(",").find((link) => link.includes('rel="next"'));
    let nextPageUrl = nextLink?.match(/<(.+)>/)[1];
    return nextPageUrl;
  }
}

/**
 * Generic service that calls the fetch API on a specified URL.
 * 
 * Adds paging and item transform capabilities.
 */
export class UrlFetchService<TRaw, TTransformed> implements PagedItemsService<TTransformed>, ItemsService<TTransformed> {
  url: string;
  options?: RequestInit;
  nextPageUrl?: string;
  itemsExtractor: ItemsExtractor<Response, TRaw>;
  itemsTransformer?: ItemsTransformer<TRaw, TTransformed>;
  totalItemsExtractor?: TotalItemsExtractor<TRaw>;
  nextPageExtractor?: NextPageUrlExtractor<Response>;

  constructor(
    {
      url,
      init,
      itemsExtractor = extractItemsFromJsonResponse,
      itemsTransformer,
      totalItemsExtractor,
      nextPageExtractor = new NextPageFromLinkResponseHeader(url)
    }: UrlFetchServiceParameters<TRaw, TTransformed>
  ) {
    this.url = url;
    this.options = init;
    this.nextPageUrl = nextPageExtractor?.firstPageLink() ?? url;
    this.itemsExtractor = itemsExtractor;
    this.itemsTransformer = itemsTransformer;
    this.totalItemsExtractor = totalItemsExtractor;
    this.nextPageExtractor = nextPageExtractor;
  }

  hasNextPage(): boolean { return !!this.nextPageUrl; }
  async getNextPageAsync(): Promise<TTransformed> {
    const url = this.nextPageUrl;
    if (!url) {
      // no url
      return null;
    }
    const response = await fetch(url, this.options);
    if (!response.ok) {
      throw new Error(`Failed to fetch from URL ${url}: ${response?.statusText}`);
    }
    var jsonResult = await this.itemsExtractor(response);
    var transformed = this.itemsTransformer ? this.itemsTransformer(jsonResult) : jsonResult as any;
    var totalItems = this.totalItemsExtractor?.call(null, jsonResult);
    // check if response has next page
    this.nextPageUrl = this.nextPageExtractor?.nextPageLink({ response, totalItems });
    return transformed;
  }
  async processAllAsync({ processor }: ProcessArgs<TTransformed>): Promise<void> {
    while (this.hasNextPage()) {
      const page = await this.getNextPageAsync();
      let res = processor(page);

      // Await if async processor.
      if (typeof ((res as any)?.then) === 'function') {
        await res;
      }
    }
  }
}
