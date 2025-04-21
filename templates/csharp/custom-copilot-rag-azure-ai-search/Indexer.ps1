param (
    [Parameter(Mandatory = $true)]
    [string]$run
)

# check if the run command is valid
$run = $run.ToLower()
if ($run -ne 'create' -and $run -ne 'delete') {
    Write-Error "Invalid run command. Please use 'create' or 'delete'"
    exit
}

$indexName = "my-documents"

# read json file
$configPath = 'appsettings.Playground.json'
$config = Get-Content -Path $configPath -Raw | ConvertFrom-Json

# load the config
$OPENAI_API_KEY = $config.OpenAI.ApiKey
$AZURE_OPENAI_API_KEY = $config.Azure.OpenAIApiKey
$AZURE_OPENAI_ENDPOINT = $config.Azure.OpenAIEndpoint
$AZURE_OPENAI_EMBEDDING_DEPLOYMENT_NAME = $config.Azure.OpenAIEmbeddingDeploymentName
$AI_SEARCH_API_KEY = $config.Azure.AISearchApiKey
$AI_SEARCH_ENDPOINT = $config.Azure.AISearchEndpoint

# check if the required keys are provided
if (!$AI_SEARCH_API_KEY -or !$AI_SEARCH_ENDPOINT) {
    Write-Error "Azure Search API Key or Endpoint is not provided in the $configPath file."
    exit
}
# check if OpenAI is enabled
if ($config.OpenAI) {
    if (!$OPENAI_API_KEY) {
        Write-Error "OpenAI API Key is not provided in the $configPath file."
        exit
    }
}
else {
    # check if Azure OpenAI is enabled
    if (!$AZURE_OPENAI_API_KEY -or !$AZURE_OPENAI_ENDPOINT -or !$AZURE_OPENAI_EMBEDDING_DEPLOYMENT_NAME) {
        Write-Error "Azure OpenAI API Key, Endpoint, or Deployment Name is not provided in the $configPath file."
        exit
    }
}

function CreateAuzreAISearchIndex {
    param (
        [string]$indexName
    )
    try {
        $indexSchema = @{
            "name"         = $indexName
            "fields"       = @(
                @{
                    "name"       = "DocId"
                    "type"       = "Edm.String"
                    "key"        = $true
                    "filterable" = $true
                    "sortable"   = $true
                },
                @{
                    "name"       = "DocTitle"
                    "type"       = "Edm.String"
                    "searchable" = $true
                    "filterable" = $true
                    "sortable"   = $true
                },
                @{
                    "name"       = "Description"
                    "type"       = "Edm.String"
                    "searchable" = $true
                    "analyzer"   = "en.lucene"
                },
                @{
                    "name"                = "DescriptionVector"
                    "type"                = "Collection(Edm.Single)"
                    "searchable"          = $true
                    "dimensions"          = 1536
                    "vectorSearchProfile" = "my-vector-config"
                    "retrievable"         = $true
                }
            )
            "corsOptions"  = @{
                "allowedOrigins" = @("*")
            }
            "vectorSearch" = @{
                "algorithms" = @(
                    @{
                        "name" = "vector-search-algorithm"
                        "kind" = "hnsw"
                    }
                )
                "profiles"   = @(
                    @{
                        "name"      = "my-vector-config"
                        "algorithm" = "vector-search-algorithm"
                    }
                )
            }
        }
        # Convert the index schema to JSON
        $indexSchemaJson = $indexSchema | ConvertTo-Json -Depth 10

        # Create the search index
        $uri = "$AI_SEARCH_ENDPOINT/indexes('$indexName')?api-version=2024-07-01"
        $headers = @{
            "Content-Type" = "application/json"
            "api-key"      = $AI_SEARCH_API_KEY
        }

        Invoke-RestMethod -Uri $uri -Method Put -Headers $headers -Body $indexSchemaJson
        # Wait for 5 seconds to allow the index to be created
        Start-Sleep -Seconds 5
    }
    catch {
        <#Do this if a terminating exception happens#>
        Write-Error "Failed to create the search index"
        throw $_
    }
}

function DeleteAuzreAISearchIndex {
    param (
        [string]$indexName
    )
    try {
        $uri = "$AI_SEARCH_ENDPOINT/indexes('$indexName')?api-version=2024-07-01"
        $headers = @{
            "api-key" = $AI_SEARCH_API_KEY
        }

        Invoke-RestMethod -Uri $uri -Method Delete -Headers $headers
    }
    catch {
        <#Do this if a terminating exception happens#>
        Write-Error "Failed to delete the search index"
        throw $_
    }
}

function GetAzureAISearchIndex {
    param (
        [string]$indexName
    )
    try {
        # Define the URI for the request
        $uri = "$AI_SEARCH_ENDPOINT/indexes('$indexName')?api-version=2024-07-01"

        # Define the headers for the request
        $headers = @{
            "Content-Type" = "application/json"
            "api-key"      = $AI_SEARCH_API_KEY
        }

        # Send the GET request to retrieve the indexes
        Invoke-RestMethod -Uri $uri -Method Get -Headers $headers
        return $true
    }
    catch {
        <#Do this if a terminating exception happens#>
        $StatusCode = $_.Exception.Response.StatusCode
        if ($StatusCode -eq 'NotFound' ) {
            return $false
        }
        Write-Error "Failed to get the search index"
        throw $_
    }
}

# function UploadDocument 
function UploadDocuments {
    param (
        $indexName,
        $documents
    )
    try {
        $body = @{
            "value" = $documents
        }
        $bodyJson = $body | ConvertTo-Json -Depth 10

        # Define the URI for the request
        $uri = "$AI_SEARCH_ENDPOINT/indexes('$indexName')/docs/search.index?api-version=2024-07-01"

        # Define the headers for the request
        $headers = @{
            "Content-Type" = "application/json"
            "api-key"      = $AI_SEARCH_API_KEY
        }
        # Send the POST request to update the document
        Invoke-RestMethod -Uri $uri -Method Post -Headers $headers -Body $bodyJson
    }
    catch {
        <#Do this if a terminating exception happens#>
        Write-Host "Failed to upload the documents"
        throw $_
    }
}
# write upload
function GetEmbeddings {
    param (
        [string]$text
    )

    if ($OPENAI_API_KEY) {
        $headers = @{
            "Content-Type"  = "application/json"
            "Authorization" = "Bearer $OPENAI_API_KEY"
        } 
    
        $body = [ordered]@{
            input = $text
            model  = "text-embedding-ada-002"
        } | ConvertTo-Json
    
        $url = "https://api.openai.com/v1/embeddings"
    }
    else {
        $openai_api_version = '2024-02-01'

        $headers = [ordered]@{
            'api-key' = $AZURE_OPENAI_API_KEY
        }
    
        $body = [ordered]@{
            input = $text
        } | ConvertTo-Json
    
        $url = "$($AZURE_OPENAI_ENDPOINT)/openai/deployments/$($AZURE_OPENAI_EMBEDDING_DEPLOYMENT_NAME)/embeddings?api-version=$($openai_api_version)"
    }

    $response = Invoke-RestMethod -Uri $url -Headers $headers -Body $body -Method Post -ContentType 'application/json'
    return $response.data[0].embedding
}

function GetData {
    $documents = @()
    $folderPath = "data"

    # Get all files in the specified folder
    $files = Get-ChildItem -Path $folderPath

    for ($i = 0; $i -lt $files.Length; $i++) {
        # Write-Output "Index: $i, Value: $($array[$i])"
        $file = $files[$i]
        # Read the content of the file
        $content = Get-Content -Path $file.FullName -Raw
        # Print the content of the file
        # Write-Output $content
        $vector = GetEmbeddings -text $content
        $document = @{
            "DocId"             = ($i + 1).ToString()
            "DocTitle"          = $file.Name
            "Description"       = $content.ToString()
            "DescriptionVector" = $vector
            "@search.action"    = "mergeOrUpload"
        }
        $documents += $document
    }
    return $documents
}
function Create {
    $exist = GetAzureAISearchIndex -indexName $indexName
    if (!$exist) {
        Write-Host "Creating index $indexName"
        CreateAuzreAISearchIndex -indexName $indexName
    }
    Write-Host "Preparing to upload documents to index $indexName"
    $documents = GetData
    Write-Host "Uploading documents to index $indexName"
    UploadDocuments -indexName $indexName -documents $documents
}

function Main {
    if ($run -eq 'create') {
        Create
        Write-Host "Index $indexName created successfully"
    }
    elseif ($run -eq 'delete') {
        DeleteAuzreAISearchIndex -indexName $indexName
        Write-Host "Index $indexName deleted successfully"
    }    
}

Main