from dataclasses import dataclass
from typing import Optional, List
from azure.search.documents.indexes.models import _edm as EDM
from azure.search.documents.models import VectorQuery, VectorizedQuery
{{#useAzureOpenAI}}
from teams.ai.embeddings import AzureOpenAIEmbeddings, AzureOpenAIEmbeddingsOptions
{{/useAzureOpenAI}}
{{#useOpenAI}}
from teams.ai.embeddings import OpenAIEmbeddings, OpenAIEmbeddingsOptions
{{/useOpenAI}}

from config import Config

async def get_embedding_vector(text: str):
    {{#useAzureOpenAI}}
    client = OpenAI(
        api_key=Config.AZURE_OPENAI_API_KEY,
        api_base=Config.AZURE_OPENAI_ENDPOINT,
        api_version=2024-10-21
    )
    {{/useAzureOpenAI}}
    {{#useOpenAI}}
    client = OpenAI(api_key=Config.OPENAI_API_KEY)
    {{/useOpenAI}}
    
    result = await client.embeddings.create(model=Config.OPENAI_EMBEDDING_DEPLOYMENT, input=query)
    if (result.status != 'success' or not result.data):
        raise Exception(f"Failed to generate embeddings for description: {text}")
    
    return result.data[0].embedding

@dataclass
class Doc:
    docId: Optional[str] = None
    docTitle: Optional[str] = None
    description: Optional[str] = None
    descriptionVector: Optional[List[float]] = None

@dataclass
class AzureAISearchDataSourceOptions:
    name: str
    indexName: str
    azureAISearchApiKey: str
    azureAISearchEndpoint: str

from azure.core.credentials import AzureKeyCredential
from azure.search.documents import SearchClient
import json

@dataclass
class Result:
    def __init__(self, output):
        self.output = output

class AzureAISearchDataSource():
    def __init__(self, options: AzureAISearchDataSourceOptions):
        self.name = options.name
        self.options = options
        self.searchClient = SearchClient(
            options.azureAISearchEndpoint,
            options.indexName,
            AzureKeyCredential(options.azureAISearchApiKey)
        )
        
    def name(self):
        return self.name

    async def render_data(self, query):
        embedding = await get_embedding_vector(query)
        vector_query = VectorizedQuery(vector=embedding, k_nearest_neighbors=2, fields="descriptionVector")

        if not query:
            return Result('')

        selectedFields = [
            'docTitle',
            'description',
            'descriptionVector',
        ]

        searchResults = self.searchClient.search(
            search_text=query,
            select=selectedFields,
            vector_queries=[vector_query],
        )

        if not searchResults:
            return Result('')


        return Result(doc)