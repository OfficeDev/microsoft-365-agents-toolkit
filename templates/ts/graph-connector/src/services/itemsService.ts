export interface ProcessArgs<T> {
  /**
   * A function that is called with items from the data source
   */
  processor: ItemsProcessorMaybeAsync<T>
}
/**
 * An interface for fetching data from a data source.
 */
export interface ItemsService<T> {
  /**
   * Processes all items from the data source.
   * 
   * @param args Process arguments.
   */
  processAllAsync(args: ProcessArgs<T>): Promise<void>;
}

/**
 * An interface for fetching data from a data source in pages.
 */
export interface PagedItemsService<T> {
  /**
   * Returns true if the service has more items available.
   */
  hasNextPage(): boolean;
  /**
   * Returns the next page of items from the service.
   * 
   * If there's no next page, an empty list should be returned.
   */
  getNextPageAsync(): Promise<T>;
}

export interface ResponseBasedPagingParameters<TResponse> {
  response: TResponse;
  totalItems?: number;
}

export type PagingParameters<TResponse> = ResponseBasedPagingParameters<TResponse>;

/**
 * Base interface for transformer functions
 */
interface Transformer<A, B> {
  (source: A): B;
}

/**
 * A function that is used to process items from a data source.
 */
export type ItemsProcessor<T> = Transformer<T, Promise<void>>;
/**
 * A function that is used to process items from a data source synchronously.
 */
export type ItemsProcessorSync<T> = Transformer<T, void>;
export type ItemsProcessorMaybeAsync<T> = ItemsProcessor<T> | ItemsProcessorSync<T>;
/**
 * Function that gets the items from an API response.
 */
export type ItemsExtractor<A, B> = Transformer<A, Promise<B>>;
export type ItemsTransformer<A, B> = Transformer<A, B>;
export type TotalItemsExtractor<A> = Transformer<A, number | null | undefined>;

export type NextPageLinkResult = string | null | undefined;
/**
 * Interface for getting the next page's URL.
 * 
 * Should return null or undefined if there's no next page.
 */
export interface NextPageUrlExtractor<A> {
  firstPageLink(): NextPageLinkResult;
  /**
   * Gets the next page's link.
   * 
   * @param params The paging parameters.
   * @returns The next page's URL or null if there's no next page.
   */
  nextPageLink(params: PagingParameters<A>): NextPageLinkResult;
}
