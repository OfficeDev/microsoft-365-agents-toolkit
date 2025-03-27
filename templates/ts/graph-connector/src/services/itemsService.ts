export interface ProcessArgs<T> {
  /**
   * Whether to disable batching during processing.
   * 
   * Disabling batching will increase memory consumption while processing items.
   * 
   * When false, the process function is allowed to repeatedly call processor as new data becomes available
   * When true, the process function must only call the processor once with all the items in the data source
   */
  disableBatching?: boolean
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
  getNextPageAsync(): Promise<T[]>;
}

/**
 * Base interface for transformer functions
 */
interface Transformer<A, B> {
  (source: A): B;
}

/**
 * A function that is used to process items from a data source.
 */
export type ItemsProcessor<T> = Transformer<T[], Promise<void>>;
/**
 * A function that is used to process items from a data source synchronously.
 */
export type ItemsProcessorSync<T> = Transformer<T[], void>;
export type ItemsProcessorMaybeAsync<T> = ItemsProcessor<T> | ItemsProcessorSync<T>;
/**
 * Function that gets the items from an API response.
 */
export type ItemsExtractor<A, B> = Transformer<A, Promise<B>>;
/**
 * Function that gets the items from an API response synchronously.
 */
export type ItemsExtractorSync<A, B> = Transformer<A, B>;
export type ItemsExtractorMaybeAsync<A, B> = ItemsExtractor<A, B> | ItemsExtractorSync<A, B>;

type NextPageResult = string | null | undefined;
/**
 * Function that gets the next page's URL.
 * 
 * Should return null or undefined if there's no next page.
 */
export type NextPageUrlExtractor<A> = Transformer<A, Promise<NextPageResult>>;
/**
 * Function that gets the next page's URL synchronously.
 * 
 * Should return null or undefined if there's no next page.
 */
export type NextPageUrlExtractorSync<A> = Transformer<A, NextPageResult>;
export type NextPageUrlExtractorMaybeAsync<A> = NextPageUrlExtractor<A> | NextPageUrlExtractorSync<A>;
