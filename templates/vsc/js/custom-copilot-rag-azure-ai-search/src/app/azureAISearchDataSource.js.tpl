const { OpenAIEmbeddings } = require("@microsoft/teams-ai");
const { AzureKeyCredential, SearchClient } = require("@azure/search-documents");

/**
 * A data source that searches through Azure AI search.
 */
class AzureAISearchDataSource {
    /**
     * Creates a new `AzureAISearchDataSource` instance.
     */
    constructor(options) {
        this.name = options.name;
        this.options = options;
        this.searchClient = new SearchClient(
            options.azureAISearchEndpoint,
            options.indexName,
            new AzureKeyCredential(options.azureAISearchApiKey),
            {}
        );
    }

    /**
     * Renders the data source as a string of text.
     */
    async renderData(context, memory, tokenizer, maxTokens) {
        const query = memory.getValue("temp.input");
        if(!query) {
            return { output: "", length: 0, tooLong: false };
        }
        
        const selectedFields = [
            "docId",
            "docTitle",
            "description",
        ];

        // hybrid search
        const queryVector= await this.getEmbeddingVector(query);
        const searchResults = await this.searchClient.search(query, {
            searchFields: ["docTitle", "description"],
            select: selectedFields,
            vectorSearchOptions: {
                queries: [
                    {
                        kind: "vector",
                        fields: ["descriptionVector"],
                        kNearestNeighborsCount: 2,
                        // The query vector is the embedding of the user's input
                        vector: queryVector
                    }
                ]
            },
        });

        if (!searchResults.results) {
            return { output: "", length: 0, tooLong: false };
        }

        // Concatenate the documents string into a single document
        // until the maximum token limit is reached. This can be specified in the prompt template.
        let usedTokens = 0;
        let doc = "";
        for await (const result of searchResults.results) {
            const formattedResult = this.formatDocument(`${result.document.description}\n Citation title:${result.document.docTitle}.`);
            const tokens = tokenizer.encode(formattedResult).length;

            if (usedTokens + tokens > maxTokens) {
                break;
            }

            doc += formattedResult;
            usedTokens += tokens;
        }

        return { output: doc, length: usedTokens, tooLong: usedTokens > maxTokens };
    }

    /**
     * Formats the result string 
     */
    formatDocument(result) {
        return `<context>${result}</context>`;
    }

    /**
     * Generate embeddings for the user's input.
     */
    async getEmbeddingVector(text) {
        {{#useOpenAI}}
        const embeddings = new OpenAIEmbeddings({
            apiKey: this.options.apiKey,
            model: this.options.openAIEmbeddingModelName,
        });
        const result = await embeddings.createEmbeddings(this.options.openAIEmbeddingModelName, text);
        {{/useOpenAI}}
        {{#useAzureOpenAI}}
        const embeddings = new OpenAIEmbeddings({
            azureApiKey: this.options.azureOpenAIApiKey,
            azureEndpoint: this.options.azureOpenAIEndpoint,
            azureDeployment: this.options.azureOpenAIEmbeddingDeploymentName,
        });

        const result = await embeddings.createEmbeddings(this.options.azureOpenAIEmbeddingDeploymentName, text);
        {{/useAzureOpenAI}}

        if (result.status !== "success" || !result.output) {
            throw new Error(`Failed to generate embeddings for description: ${text}`);
        }

        return result.output[0];
    }
}

module.exports = {
  AzureAISearchDataSource,
};