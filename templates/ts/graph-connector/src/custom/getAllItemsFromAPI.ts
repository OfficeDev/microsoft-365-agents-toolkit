import { Config } from "../models/Config";
import { Item } from "../models/Item";

async function getPaginatedIssues(config: Config, per_page: number, repo: string, since?: Date) {
  const paginatedResponse = [];
  // since = new Date("2020-10-01T00:00:00Z");
  // console.log(since.toISOString());
  const response = await fetch(
    `https://api.github.com/repos/${repo}/issues?state=all&per_page=${per_page}${
      since ? `&since=${since.toISOString()}` : ""
    }`,
    {
      headers: {
        Authorization: `Bearer ${config.connector.accessToken}`,
      },
    }
  );

  if (!response.ok) {
    throw new Error(`Failed to fetch items in repo ${repo}: ${response.statusText}`);
  }

  paginatedResponse.push(await response.json());

  if (response.headers.get("link")) {
    // get next page of results
    let nextPage = response.headers
      .get("link")
      .split(",")
      .find((link) => link.includes('rel="next"'));
    while (nextPage) {
      // get next page url
      const nextPageUrl = nextPage.match(/<(.+)>/)[1];
      const response = await fetch(nextPageUrl, {
        headers: {
          Authorization: `Bearer ${config.connector.accessToken}`,
        },
      });
      paginatedResponse.push(await response.json());
      nextPage = response.headers
        .get("link")
        .split(",")
        .find((link) => link.includes('rel="next"'));
    }
  }

  return paginatedResponse.flat();
}
/**
 * Gets all items from the repository.
 * @param config - The configuration object.
 * @returns An array of items.
 */
export async function getAllItemsFromAPI(config: Config, since?: Date): Promise<Item[]> {
  const repos = config.connector.repos.split(",");

  const items = await Promise.all(
    repos.map(async (repo) => {
      const issues = await getPaginatedIssues(config, 100, repo, since);
      return issues.filter((issue) => !issue.pull_request);
    })
  );

  return (
    items.length > 0 &&
    items.flat().map<Item>((issue) => {
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
      };
    })
  );
}
