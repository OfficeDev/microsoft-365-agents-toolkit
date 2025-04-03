import { Config } from "../models/Config";
import { Item } from "../models/Item";
import { ItemsExtractor, ItemsService, NextPageUrlExtractor, ProcessArgs } from "../services/itemsService";
import { UrlFetchService } from "./urlFetchService";

/**
 * Options used to construct the GitHub issues fetch service.
 */
interface GitHubIssuesFetchParameters {
  config: Config;
  pageSize?: number;
  since?: Date;
  itemsExtractor?: ItemsExtractor<Response, Item[]>;
  nextPageExtractor?: NextPageUrlExtractor<Response>;
}

export class MultiRepoIssuesFetchService implements ItemsService<Item[]> {
  services: ItemsService<Item[]>[];

  constructor(services: ItemsService<Item[]>[]) {
    this.services = services;
  }

  static forRepositories(
    repositories: string[],
    options: GitHubIssuesFetchParameters,
  ): MultiRepoIssuesFetchService {
    if (repositories.length < 1) {
      throw Error('repositories must have at least one element');
    }
    let services = repositories.map(repo => {
      const url = buildGitHubUrl(repo, options.pageSize, options.since);
      const init: RequestInit = {};
      if (options.config.connector.accessToken) {
        init.headers = { Authorization: `Bearer ${options.config.connector.accessToken}` }
      }
      return new UrlFetchService({
        url, init, itemsTransformer: githubItemsTransformer,
      });
    });
    return new MultiRepoIssuesFetchService(services);
  }

  async processAllAsync({ processor }: ProcessArgs<Item[]>): Promise<void> {
    await Promise.all(this.services.map(async svc => {
      await svc.processAllAsync({ processor })
    }));
  }
}

function githubItemsTransformer(items: Item[]): Item[] {
  const issues = items.filter((issue: any) => !issue.pull_request);

  return issues.map((issue: any) => {
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

function buildGitHubUrl(repo: string, pageSize?: number, since?: Date) : string {
  const perPage = pageSize ? `&per_page=${pageSize}` : '';
  const sinceStr = since ? `&since=${since.toISOString()}` : '';
  return `https://api.github.com/repos/${repo}/issues?state=all${perPage}${sinceStr}`;
}
