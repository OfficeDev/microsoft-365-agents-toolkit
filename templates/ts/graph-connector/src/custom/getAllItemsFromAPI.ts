import { Config } from "../models/Config";
import { Item } from "../models/Item";

/**
 * Fetches paginated issues from the GitHub API.
 * @param config - The configuration object.
 * @param per_page - The number of items per page.
 * @param repo - The repository name in the format 'owner/repo'.
 * @param since - Optional date to filter issues updated after this date.
 * @returns An array of issues.
 */
async function getPaginatedIssues(config: Config, per_page: number, repo: string, since?: Date) {
  const paginatedResponse = [];
  const fetchUrl = `https://api.github.com/repos/${repo}/issues?state=all&per_page=${per_page}${
    since ? `&since=${since.toISOString()}` : ""
  }`;
  const response = await fetchIssues(config, fetchUrl, repo);
  paginatedResponse.push(await response.json());

  let nextPageUrl = getNextPageUrl(response);
  while (nextPageUrl) {
    const response = await fetchIssues(config, nextPageUrl, repo);
    paginatedResponse.push(await response.json());
    
    nextPageUrl = getNextPageUrl(response);
  }

  return paginatedResponse.flat();
}

/**
  * Fetches issues from the GitHub API.
  * @param config - The configuration object.
  * @param fetchUrl - The URL to fetch issues from.
  * @param repo - The repository name in the format 'owner/repo'.
  * @returns A promise that resolves to the response.
  */
async function fetchIssues(config: Config, fetchUrl: string, repo: string): Promise<Response> {
  // Use the access token from the config if available
  // to authenticate the request to the GitHub API if using a private repo
  // or to avoid rate limiting
  const headers: HeadersInit = {};
  if (config.connector.accessToken) {
    headers["Authorization"] = `Bearer ${config.connector.accessToken}`;
  }
  const response = await fetch(fetchUrl, {
    headers,
  });
  if (!response.ok) {
    throw new Error(`Failed to fetch items in repo ${repo}: ${response.statusText}`);
  }
  return response;
}

/**
 * Gets the URL for the next page of results from the response headers.
 * Parse the link header to find the next page URL.
 * @see https://docs.github.com/en/rest/using-the-rest-api/using-pagination-in-the-rest-api?apiVersion=2022-11-28#using-link-headers
 * @param response - The response object from the fetch request.
 * @returns The URL for the next page of results, or null if there are no more pages.
 */
function getNextPageUrl(response: Response): string | null {
  const linkHeader = response.headers.get("link");
  if (!linkHeader) {
    return null;
  }
  const links = linkHeader.split(", ");
  for (const link of links) {
    const match = link.match(/<([^>]+)>; rel="next"/);
    if (match) {
      return match[1];
    }
  }
  return null;
}

// [Customization point]
// If you need additional logic to get all items from the repository, you can add it here
// This function is used to get all items from the repository.
// The items are filtered to exclude pull requests and only include issues.
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
