# read json file
$configPath = 'appsettings.Playground.json'
$config = Get-Content -Path $configPath -Raw | ConvertFrom-Json

# load the config
$OPENAI_API_KEY = $config.OpenAI.ApiKey
$AZURE_OPENAI_API_KEY = $config.Azure.OpenAIApiKey
$AZURE_OPENAI_ENDPOINT = $config.Azure.OpenAIEndpoint
$Auzre_OpenAI_DEPLOYMENTNAME = $config.Azure.OpenAIDeploymentName

# check if OpenAI is enabled
if ($config.OpenAI) {
    if (!$OPENAI_API_KEY) {
        Write-Error "OpenAI API Key is not provided in the $configPath file."
        exit    
    }
}
else {
    # check if Azure OpenAI is enabled
    if (!$AZURE_OPENAI_API_KEY -or !$AZURE_OPENAI_ENDPOINT -or !$Auzre_OpenAI_DEPLOYMENTNAME) {
        Write-Error "Azure OpenAI API Key, Endpoint, or Deployment Name is not provided in the $configPath file."
        exit
    }
}

function Create {
    $functionGetCurrentWeather = @{
        type     = "function"
        function = @{
            name        = "getCurrentWeather"
            description = "Get the weather in location"
            parameters  = @{
                type       = "object"
                properties = @{
                    location = @{
                        type        = "string"
                        description = "The city and state e.g. San Francisco, CA"
                    }
                    unit     = @{
                        type = "string"
                        enum = @("c", "f")
                    }
                }
                required   = @("location")
            }
        }
    }
    $functionGetNickName = @{
        type     = "function"
        function = @{
            name        = "getNickname"
            description = "Get the nickname of a city"
            parameters  = @{
                type       = "object"
                properties = @{
                    location = @{
                        type        = "string"
                        description = "The city and state e.g. San Francisco, CA"
                    }
                }
                required   = @("location")
            }
        }
    }
    $tools = @(@{type = "code_interpreter" }, $functionGetCurrentWeather, $functionGetNickName)

    try {
        if ($OPENAI_API_KEY) {
            $headers = @{
                "Content-Type"  = "application/json"
                "Authorization" = "Bearer $OPENAI_API_KEY"
                "OpenAI-Beta"   = "assistants=v2"
            }
    
            $body = @{
                instructions = "You are an intelligent agent that can\n - write and run code to answer math questions\n - use the provided functions to answer questions"
                name         = "Assistant"
                tools        = $tools
                model        = "gpt-3.5-turbo"
            } | ConvertTo-Json -Depth 10
    
            $uri = "https://api.openai.com/v1/assistants"
        }
        else {
            $headers = @{
                "Content-Type" = "application/json"
                "api-key"      = $AZURE_OPENAI_API_KEY
            }
        
            $body = @{
                instructions = "You are an intelligent agent that can\n - write and run code to answer math questions\n - use the provided functions to answer questions"
                name         = "Assistant"
                tools        = $tools
                model        = $Auzre_OpenAI_DEPLOYMENTNAME
            } | ConvertTo-Json -Depth 10
            
            $uri = "$AZURE_OPENAI_ENDPOINT/openai/assistants?api-version=2024-05-01-preview"
        }
    
        $res = Invoke-RestMethod -Uri $uri -Method Post -Body $body -Headers $headers
        Write-Host "Create assistant completed"
        Write-Host "The assistant id is $($res.id)"
    }
    catch {
        Write-Host "Failed to create the assistant"
        throw $_
    }
}

Create