using Microsoft.Bot.Builder;
using Microsoft.Teams.AI.AI;
using Microsoft.Teams.AI.AI.Action;

namespace {{SafeProjectName}}
{
    public class ActionHandlers
    {
        public class WeatherParameters
        {
            public string Location { get; set; }
            public string Unit { get; set; }
        }

        [Action(AIConstants.HttpErrorActionName)]
        public async Task<string> OnHttpError([ActionTurnContext] ITurnContext turnContext)
        {
            await turnContext.SendActivityAsync("An AI request failed. Please try again later.");
            return AIConstants.StopCommand;
        }
        Dictionary<string, Dictionary<string, string>> WeatherData = new Dictionary<string, Dictionary<string, string>>
        {
            { "San Francisco", new Dictionary<string, string> { { "f", "71.6F" }, { "c", "22C" } } },
            { "Los Angeles", new Dictionary<string, string> { { "f", "75.2F" }, { "c", "24C" } } }
        };
        [Action("getCurrentWeather")]
        public async Task<string> OnGetCurrentWeather([ActionTurnContext] ITurnContext turnContext, [ActionParameters] Dictionary<string, object> entities)
        {
            if (entities.TryGetValue("location", out object locationObj))
            {
                string location = locationObj.ToString();
                if (location.Contains("San Francisco") || location.Contains("Los Angeles"))
                {
                    string unit = "f";
                    if (entities.TryGetValue("unit", out object unitObj))
                    {
                        unit = unitObj.ToString();
                    }
                    if (location.Contains("San Francisco"))
                    {
                        return WeatherData["San Francisco"][unit];
                    }
                    else
                    {
                        return WeatherData["Los Angeles"][unit];
                    }
                }
                else
                {
                    return $"No weather data for {location} found";
                }
            }
            else
            {
                return "No location is found in parameters";
            }
        }
        Dictionary<string, string> NicknameData = new Dictionary<string, string>
        {
            { "San Francisco", "The Golden City" },
            { "Los Angeles", "LA" }
        };
        [Action("getNickname")]
        public async Task<string> OnGetNickname([ActionTurnContext] ITurnContext turnContext, [ActionParameters] Dictionary<string, object> entities)
        {
            if (entities.TryGetValue("location", out object locationObj))
            {
                string location = locationObj.ToString();
                if (location.Contains("San Francisco"))
                {
                    return NicknameData["San Francisco"];
                }
                else if (location.Contains("Los Angeles"))
                {
                    return NicknameData["Los Angeles"];
                }
                else
                {
                    return $"No nickname data for {location} found";
                }
            }
            else
            {
                return "No location is found in parameters";
            }
        }
    }
}
