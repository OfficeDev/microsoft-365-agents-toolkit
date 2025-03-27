import { Config } from "../models/Config";
import { Item } from "../models/Item";
import { ItemsExtractorMaybeAsync, ItemsService, NextPageUrlExtractorMaybeAsync, PagedItemsService, ProcessArgs } from "../services/itemsService";
import { asPromise } from "../utils";
import { UrlFetchService } from "./urlFetchService";

/**
 * Options used to construct the GitHub issues fetch service.
 */
interface GitHubIssuesFetchServiceParameters {
  config: Config;
  pageSize?: number;
  since?: Date;
  itemsExtractor?: ItemsExtractorMaybeAsync<Response, Item[]>;
  nextPageExtractor?: NextPageUrlExtractorMaybeAsync<Response>;
}

/**
 * Fetches GitHub issues from a GitHub repository.
 * 
 * Uses the generic URL fetch service.
 */
export class GitHubIssuesFetchService implements PagedItemsService<Item>, ItemsService<Item> {
  repo: string;
  pageSize: number;
  since: Date | null;
  config: Config;
  urlFetchSvc: UrlFetchService<Item>;

  constructor(
    repo: string,
    {
      config,
      pageSize = 100,
      since,
      itemsExtractor = gitHubExtractItemsFromResponse,
      nextPageExtractor = gitHubNextPageFromResponseSync
    }: GitHubIssuesFetchServiceParameters
  ) {
    this.config = config;
    this.repo = repo;
    this.pageSize = pageSize;
    this.since = since;

    const sinceStr = since ? `&since=${since.toISOString()}` : ""
    const url = `https://api.github.com/repos/${repo}/issues?state=all&per_page=${this.pageSize}${sinceStr}`;
    const init: RequestInit = {};
    if (config.connector.accessToken) {
      init.headers = { Authorization: `Bearer ${this.config.connector.accessToken}` }
    }
    this.urlFetchSvc = new UrlFetchService({
      url, init, itemsExtractor, nextPageExtractor
    });
  }

  hasNextPage(): boolean { return this.urlFetchSvc.hasNextPage(); }

  async getNextPageAsync(): Promise<Item[]> {
    const result = (await this.urlFetchSvc.getNextPageAsync()).filter((issue: any) => !issue.pull_request);
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

  async processAllAsync({ disableBatching, processor }: ProcessArgs<Item>): Promise<void> {
    if (!disableBatching) {
      while (this.hasNextPage()) {
        var page = await this.getNextPageAsync();
        await asPromise(processor(page));
      }
    } else {
      const allItems = [];
      while (this.hasNextPage()) {
        var page = await this.getNextPageAsync();
        allItems.push(...page);
      }
      await asPromise(processor(allItems));
    }
  }
}

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
export function gitHubNextPageFromResponseSync(response: Response): string | null {
  let nextLink = response.headers.get("link").split(",").find((link) => link.includes('rel="next"'));
  let nextPageUrl = nextLink?.match(/<(.+)>/)[1];
  return nextPageUrl
}

export class MultiRepoIssuesFetchService implements ItemsService<Item> {
  services: ItemsService<Item>[];

  constructor(services: ItemsService<Item>[]) {
    this.services = services;
  }

  static forRepositories(
    repositories: string[],
    options: GitHubIssuesFetchServiceParameters,
  ): MultiRepoIssuesFetchService {
    if (repositories.length < 1) {
      throw Error('repositories must have at least one element');
    }
    let services = repositories.map(repo => new GitHubIssuesFetchService(repo, options));
    return new MultiRepoIssuesFetchService(services);
  }

  async processAllAsync({ disableBatching, processor }: ProcessArgs<Item>): Promise<void> {
    if (!disableBatching) {
      await Promise.all(this.services.map(async svc => {
        await svc.processAllAsync({ disableBatching, processor })
      }));
    } else {
      const allItems = [];
      await Promise.all(this.services.map(async svc => {
        await svc.processAllAsync({
          disableBatching,
          processor: items => {
            for (let i = 0; i < items.length; i++) {
              const it = items[i];
              allItems.push(it);
            }
          }
        })
      }));
      await processor(allItems);
    }
  }
}
