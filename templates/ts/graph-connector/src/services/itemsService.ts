
/**
 * An interface for fetching data from a data source.
 */
export interface ItemsService<T> {
  /**
   * Returns a list of all the items in a data source.
   * 
   * This function is expected to return all items in the source.
   */
  getAllAsync(): Promise<T[]>;

  /**
   * Processes all items from the data source in batches.
   * 
   * @param processor A function to process each batch of items
   */
  processAllAsync(processor: ItemsProcessor<T>): Promise<void>;
}

/**
 * An interface for fetching data from a data source in pages.
 */
export interface PagedItemsService<T> extends ItemsService<T> {
  /**
   * Returns true if the service has more items to return.
   */
  hasNextPage(): boolean;
  /**
   * Returns the next page of items from the service.
   */
  getNextPageAsync(): Promise<T[]>;
}

/**
 * Base interface for transformer functions
 */
export interface Transformer<A, B> {
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
/**
 * Function that gets the items from an API response.
 */
export type ItemsExtractor<A, B> = Transformer<A, Promise<B>>;
/**
 * Function that gets the items from an API response synchronously.
 */
export type ItemsExtractorSync<A, B> = Transformer<A, B>;

/**
 * Function that gets the next page's URL
 */
export type NextPageUrlExtractor<A> = Transformer<A, Promise<string | null>>;
/**
 * Function that gets the next page's URL synchronously
 */
export type NextPageUrlExtractorSync<A> = Transformer<A, string | null>;
