import { DataSource, Memory, RenderedPromptSection, Tokenizer } from "@microsoft/teams-ai";
import { TurnContext } from "botbuilder";
import { Client, ResponseType } from "@microsoft/microsoft-graph-client";

/**
 * A data source that searches through Graph API.
 */
export class GraphDataSource implements DataSource {
    /**
     * Name of the data source.
     */
    public readonly name: string;

    /**
     * Graph client to make requests to Graph API.
     */
    private graphClient: Client;

    /**
     * Creates a new instance of the Graph DataSource instance.
     */
    public constructor(name: string) {
        this.name = name;
    }

    /**
     * Renders the data source as a string of text.
     * @remarks
     * The returned output should be a string of text that will be injected into the prompt at render time.
     * @param context Turn context for the current turn of conversation with the user.
     * @param memory An interface for accessing state values.
     * @param tokenizer Tokenizer to use when rendering the data source.
     * @param maxTokens Maximum number of tokens allowed to be rendered.
     * @returns A promise that resolves to the rendered data source.
     */
    public async renderData(context: TurnContext, memory: Memory, tokenizer: Tokenizer, maxTokens: number): Promise<RenderedPromptSection<string>> {
        const query = memory.getValue("temp.input") as string;
        if(!query) {
            return { output: "", length: 0, tooLong: false };
        }
        if (!this.graphClient) {
            this.graphClient = Client.init({
                authProvider: (done) => {
                    done(null, (memory as any).temp.authTokens["graph"]);
                }
            });
        }
        let graphQuery = query;
        if (query.toLocaleLowerCase().includes("perksplus")) {
            graphQuery = "perksplus program";
        } else if (query.toLocaleLowerCase().includes("company") || query.toLocaleLowerCase().includes("history")) {
            graphQuery = "company history";
        } else if (query.toLocaleLowerCase().includes("northwind") || query.toLocaleLowerCase().includes("health") || query.toLocaleLowerCase().includes("plan")) {
            graphQuery = "northwind health";
        }

        const contentResults = [];
        const response = await this.graphClient.api("/search/query").post({
            requests: [
                {
                entityTypes: ["driveItem"],
                query: {
                    // Search for markdown files in the user's OneDrive and SharePoint
                    // The supported file types are listed here:
                    // https://learn.microsoft.com/sharepoint/technical-reference/default-crawled-file-name-extensions-and-parsed-file-types
                    queryString: `${graphQuery}`,
                },
                // This parameter is required only when searching with application permissions
                // https://learn.microsoft.com/graph/search-concept-searchall
                // region: "US",
                },
            ],
        });
        for (const value of response?.value ?? []) {
            for (const hitsContainer of value?.hitsContainers ?? []) {
              contentResults.push(...(hitsContainer?.hits ?? []));
            }
        }

        // Add documents until you run out of tokens
        let length = 0,
        output = "";
        for (const result of contentResults) {
            const rawContent = await this.downloadSharepointFile(
                result.resource.webUrl
            );
            if (!rawContent) {
                continue;
            }
            let doc = `${rawContent}\n Citation title:${result.resource.name}. Url:${result.resource.webUrl}\n\n`;
            let docLength = tokenizer.encode(doc).length;
            const remainingTokens = maxTokens - (length + docLength);
            if (remainingTokens <= 0) {
                break;
            }

            // Append do to output
            output += doc;
            length += docLength;
        }
        return { output: this.formatDocument(output), length: output.length, tooLong: false };
    }

    /**
     * Formats the result string 
     * @param result 
     * @returns 
     */
    private formatDocument(result: string): string {
        return `<context>${result}</context>`;
    }

    // Download the file from SharePoint
    // https://docs.microsoft.com/en-us/graph/api/driveitem-get-content
    private async downloadSharepointFile(
        contentUrl: string
    ): Promise<string | undefined> {
        const encodedUrl = this.encodeSharepointContentUrl(contentUrl);
        const fileContentResponse = await this.graphClient
            .api(`/shares/${encodedUrl}/driveItem/content`)
            .responseType(ResponseType.TEXT)
            .get();

        return fileContentResponse;
    }

    private encodeSharepointContentUrl(webUrl: string): string {
        const byteData = Buffer.from(webUrl, "utf-8");
        const base64String = byteData.toString("base64");
        return (
            "u!" + base64String.replace("=", "").replace("/", "_").replace("+", "_")
        );
    }
}