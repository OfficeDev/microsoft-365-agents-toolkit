import { Config } from "../models/Config";
import { Item } from "../models/Item";
import { ItemsExtractor, ItemsExtractorSync, ItemsProcessor, ItemsService, NextPageUrlExtractor, NextPageUrlExtractorSync, PagedItemsService } from "../services/itemsService";

/**
 * Extracts items from a GitHub API response.
 * 
 * @param response The Fetch API response object.
 * @returns A promise that resolves to the items
 */
export async function gitHubExtractItemsFromResponse(response: Response): Promise<Item[]> {
  return await response.json();
}

/**
 * Gets the next page's link from a GitHub API response.
 * 
 * @param response The Fetch API response object.
 * @returns A promise that resolves to the next page's URL or null if there's no next page.
 */
export function gitHubNextPageFromResponse(response: Response): string | null {
  let nextLink = response.headers.get("link").split(",").find((link) => link.includes('rel="next"'));
  let nextPageUrl = nextLink?.match(/<(.+)>/)[1];
  return nextPageUrl
}

/**
 * Options used by the fetch service
 */
class IssuesFetchServiceParameters {
  config: Config;
  pageSize?: number;
  since?: Date;
  baseUrl?: string;
  itemsExtractor?: ItemsExtractor<Response, Item[]> | ItemsExtractorSync<Response, Item[]>;
  nextPageExtractor?: NextPageUrlExtractor<Response> | NextPageUrlExtractorSync<Response>;
}

/**
 * Fetches GitHub issues from a GitHub repository.
 */
export class GitHubIssuesFetchService implements PagedItemsService<Item> {
  itemsExtractor: ItemsExtractor<Response, Item[]> | ItemsExtractorSync<Response, Item[]>;
  nextPageExtractor: NextPageUrlExtractor<Response> | NextPageUrlExtractorSync<Response>;
  baseUrl: string;
  repo: string;
  nextPageUrl?: string;
  pageSize: number;
  since: Date | null;
  config: Config;

  constructor(
    repo: string,
    {
      config,
      pageSize = 100,
      since,
      baseUrl = "https://api.github.com",
      itemsExtractor = gitHubExtractItemsFromResponse,
      nextPageExtractor = gitHubNextPageFromResponse
    }: IssuesFetchServiceParameters
  ) {
    this.config = config;
    this.repo = repo;
    this.pageSize = pageSize;
    this.since = since;
    this.baseUrl = baseUrl;
    this.itemsExtractor = itemsExtractor;
    this.nextPageExtractor = nextPageExtractor;
    const sinceStr = since ? `&since=${since.toISOString()}` : ""
    this.nextPageUrl = `${this.baseUrl}/repos/${repo}/issues?state=all&per_page=${this.pageSize}${sinceStr}`;
  }

  hasNextPage(): boolean { return !!this.nextPageUrl; }

  async getNextPageAsync(): Promise<Item[]> {
    const url = this.nextPageUrl;
    if (!url) {
      // no url
      return [];
    }
    let opts: any = {};
    if (this.config.connector.accessToken) {
      opts.headers = { Authorization: `Bearer ${this.config.connector.accessToken}` }
    }
    const response = await fetch(url, opts);
    if (!response.ok) {
      throw new Error(`Failed to fetch items in repo ${this.repo}: ${response.statusText}`);
    }
    // check if response has next page
    const res0: any = this.nextPageExtractor(response);
    // await if promise
    this.nextPageUrl = res0 && typeof(res0.then) === 'function' ? (await res0) : res0;
    const res1: any = this.itemsExtractor(response);
    let result = (res1 && typeof(res1.then) === 'function' ? (await res1) : res1).filter((issue: any) => !issue.pull_request);
    return result.map((issue: any) => {
      return {
        id: issue.id,
        issueNumber: issue.number as string,
        owner: issue.repository_url.split("/").slice(-2)[0],
        repo: issue.repository_url.split("/").slice(-1)[0],
        assignedTo: issue.assignees?.map((assignee) => assignee.login).join(", "),
        state: issue.state,
        lastModified: new Date(issue.updated_at).toISOString().slice(0, -5) + "Z",
        title: issue.title,
        abstract: issue.body,
        author: issue.user.login,
        content: `${issue.title} - ${issue.body}`,
        url: issue.html_url,
      }
    });
  }

  async getAllAsync(): Promise<Item[]> {
    const results = [];
    while (this.hasNextPage()) {
      results.push(...(await this.getNextPageAsync()));
    }
    return results;
  }
  
  async processAllAsync(processor: ItemsProcessor<Item>): Promise<void> {
    while (this.hasNextPage()) {
      var page = await this.getNextPageAsync();
      await processor(page);
    }
  }
}

export class MultiRepoIssuesFetchService implements ItemsService<Item> {
  services: ItemsService<Item>[];

  constructor(services: ItemsService<Item>[]) {
    this.services = services;
  }

  static forRepositories(
    repositories: string[],
    options: IssuesFetchServiceParameters,
  ): MultiRepoIssuesFetchService {
    if (repositories.length < 1) {
      throw Error('repositories must have at least one element');
    }
    let services = repositories.map(repo => new GitHubIssuesFetchService(repo, options));
    return new MultiRepoIssuesFetchService(services);
  }

  async getAllAsync(): Promise<Item[]> {
    let all = await Promise.all(this.services.map(svc => svc.getAllAsync()));
    return all.flat();
  }

  async processAllAsync(processor: ItemsProcessor<Item>): Promise<void> {
    await Promise.all(this.services.map(async svc => {
      await svc.processAllAsync(processor)
    }));
  }
}
