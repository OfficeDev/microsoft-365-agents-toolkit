const { Client, ResponseType } = require("@microsoft/microsoft-graph-client");

/**
 * A data source that searches through Graph API.
 */
class GraphDataSource {
    /**
     * Creates a new instance of the Graph DataSource instance.
     */
    constructor(name) {
        this.name = name;
    }

    /**
     * Renders the data source as a string of text.
     */
    async renderData(context, memory, tokenizer, maxTokens) {
        const query = memory.getValue("temp.input");
        if(!query) {
            return { output: "", length: 0, tooLong: false };
        }
        if (!this.graphClient) {
            this.graphClient = Client.init({
                authProvider: (done) => {
                    done(null, memory.temp.authTokens["graph"]);
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
     */
    formatDocument(result) {
        return `<context>${result}</context>`;
    }

    // Download the file from SharePoint
    // https://docs.microsoft.com/en-us/graph/api/driveitem-get-content
    async downloadSharepointFile(contentUrl) {
        const encodedUrl = this.encodeSharepointContentUrl(contentUrl);
        const fileContentResponse = await this.graphClient
            .api(`/shares/${encodedUrl}/driveItem/content`)
            .responseType(ResponseType.TEXT)
            .get();

        return fileContentResponse;
    }

    encodeSharepointContentUrl(webUrl) {
        const byteData = Buffer.from(webUrl, "utf-8");
        const base64String = byteData.toString("base64");
        return (
            "u!" + base64String.replace("=", "").replace("/", "_").replace("+", "_")
        );
    }
}

module.exports = {
  GraphDataSource,
};