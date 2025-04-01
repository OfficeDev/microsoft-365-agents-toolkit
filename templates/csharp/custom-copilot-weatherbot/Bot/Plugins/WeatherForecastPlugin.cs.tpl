using Microsoft.SemanticKernel;
using System;
using System.Threading.Tasks;

namespace {{SafeProjectName}}.Bot.Plugins;

public class WeatherForecastPlugin
{
    /// <summary>
    /// Retrieve the weather forecast for a specific date. This is a placeholder for a real implementation
    /// and currently only returns a random temperature. This would typically call a weather service API.
    /// </summary>
    /// <param name="date">The date as a parsable string</param>
    /// <param name="location">The location to get the weather for</param>
    /// <returns></returns>
    [KernelFunction]
    public Task<WeatherForecast> GetForecastForDate(string date,  string location)
    {
        return Task.FromResult(new WeatherForecast
        {
            Date = date,
            TemperatureC = Random.Shared.Next(-20, 55)
        });
    }
}
