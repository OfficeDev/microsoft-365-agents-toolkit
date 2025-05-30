/**
 * Defines the utility methods.
 */
const { KnownAnalyzerNames } = require("@azure/search-documents");
const { OpenAIEmbeddings } = require("@microsoft/teams-ai");
{{#useOpenAI}}
const config = require("../config");
{{/useOpenAI}}

/**
 * A wrapper for setTimeout that resolves a promise after timeInMs milliseconds.
 */
function delay(timeInMs) {
    return new Promise((resolve) => setTimeout(resolve, timeInMs));
}

/**
 * Deletes the index with the given name
 */
function deleteIndex(client, name) {
    return client.deleteIndex(name);
}

/**
 * Adds or updates the given documents in the index
 */
async function upsertDocuments(client, documents) {
    return await client.mergeOrUploadDocuments(documents);
}

/**
 * Creates the index with the given name
 */
async function createIndexIfNotExists(client, name) {
    const MyDocumentIndex = {
        name,
        fields: [
            {
                type: "Edm.String",
                name: "docId",
                key: true,
                filterable: true,
                sortable: true
            },
            {
                type: "Edm.String",
                name: "docTitle",
                searchable: true,
                filterable: true,
                sortable: true
            },
            {
                type: "Edm.String",
                name: "description",
                searchable: true,
                analyzerName: KnownAnalyzerNames.EnLucene
            },
            {
                type: "Collection(Edm.Single)",
                name: "descriptionVector",
                searchable: true,
                vectorSearchDimensions: 1536,
                vectorSearchProfileName: "my-vector-config"
            },
        ],
        corsOptions: {
            // for browser tests
            allowedOrigins: ["*"]
        },
        vectorSearch: {
            algorithms: [{ name: "vector-search-algorithm", kind: "hnsw" }],
            profiles: [
                {
                    name: "my-vector-config",
                    algorithmConfigurationName: "vector-search-algorithm"
                }
            ]
        }
    };

    await client.createOrUpdateIndex(MyDocumentIndex);
}

/**
 * Generate the embedding vector
 */
async function getEmbeddingVector(text) {
    {{#useOpenAI}}
    const embeddings = new OpenAIEmbeddings({
        apiKey: process.env.SECRET_OPENAI_API_KEY,
        model: config.openAIEmbeddingModelName
    });
    const result = await embeddings.createEmbeddings(config.openAIEmbeddingModelName, text);
    {{/useOpenAI}}
    {{#useAzureOpenAI}}
    const embeddings = new OpenAIEmbeddings({
        azureApiKey: process.env.SECRET_AZURE_OPENAI_API_KEY,
        azureEndpoint: process.env.AZURE_OPENAI_ENDPOINT,
        azureDeployment: process.env.AZURE_OPENAI_EMBEDDING_DEPLOYMENT_NAME,
    });

    const result = await embeddings.createEmbeddings( process.env.AZURE_OPENAI_EMBEDDING_DEPLOYMENT_NAME, text);
    {{/useAzureOpenAI}}

    if (result.status !== "success" || !result.output) {
        throw new Error(`Failed to generate embeddings for description: ${text}`);
    }

    return result.output[0];
}

module.exports = {
  deleteIndex,
  createIndexIfNotExists,
  delay,
  upsertDocuments,
  getEmbeddingVector,
};