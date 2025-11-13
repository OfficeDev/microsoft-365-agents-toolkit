using System.ComponentModel;
using System.Text.RegularExpressions;
using System.Text.Json;

namespace TravelAgent.Bot.Plugins
{
    public class DataPlugin
    {
        /// <summary>
        /// Call external flight/hotel API to get data. 
        /// </summary>
        /// <param name="date">The date as a parsable string</param>
        /// <param name="location">The location to get the travel plan</param>
        /// <param name="destination">The destination to get the travel plan</param>
        /// <param name="isFlight">Boolean indicating if flight information is needed</param>
        /// <param name="isHotel">Boolean indicating if hotel information is needed</param>
        /// <returns></returns>
        [Description("This function talks to external flight/hotel API and gets flight/hotel information based on user query.")]
        public async Task<string> GetHotelFlightDataAsync(string date, string location, string destination, bool isFlight, bool isHotel)
        {
            var flightData = string.Empty;
            var hotelData = string.Empty;

            try
            {
                if (isFlight)
                {
                    // Replace this mock implementation with real API calls

                    // flightData = await GetFlightDataAsync(date, location, destination);
                    flightData = await GetMockFlightDataAsync(date, location, destination);
                }
                if (isHotel)
                {
                    // Replace this mock implementation with real API calls

                    // hotelData = await GetHotelDataAsync(date, destination);
                    hotelData = await GetMockHotelDataAsync(date, destination);
                }

                return JsonSerializer.Serialize(new { flights = flightData, hotels = hotelData });
            }
            catch (Exception ex)
            {
                // Log error and return empty result
                Console.WriteLine($"Error in GetHotelFlightDataAsync: {ex.Message}");
                return JsonSerializer.Serialize(new { flights = flightData, hotels = hotelData });
            }
        }

        private static async Task<string> GetFlightDataAsync(string date, string location, string destination)
        {
            // For real implementation:
            // 1. Replace apiUrl with actual flight API endpoints
            // 2. Update requestHeaders with real authentication tokens/API keys
            // 3. Modify requestBody with actual search parameters

            var apiUrl = "https://api.example-travel.com/search";
            var requestHeaders = new Dictionary<string, string>
                {
                    { "Authorization", "Bearer YOUR_API_KEY_HERE" },
                    { "Content-Type", "application/json" },
                    { "Accept", "application/json" }
                };
            var requestBody = new
            {
                departure = location,
                destination,
                date
            };
            return await SendRequestAsync(apiUrl, requestHeaders, requestBody);
        }

        private static async Task<string> GetHotelDataAsync(string date, string location)
        {
            // For real implementation:
            // 1. Replace apiUrl with actual hotel API endpoints
            // 2. Update requestHeaders with real authentication tokens/API keys
            // 3. Modify requestBody with actual search parameters

            var apiUrl = "https://api.example-travel.com/hotels";
            var requestHeaders = new Dictionary<string, string>
                {
                    { "Authorization", "Bearer YOUR_API_KEY_HERE" },
                    { "Content-Type", "application/json" },
                    { "Accept", "application/json" }
                };
            var requestBody = new
            {
                location,
                date
            };
            return await SendRequestAsync(apiUrl, requestHeaders, requestBody);
        }

        private static async Task<string> SendRequestAsync(string url, Dictionary<string, string> headers, object body = null)
        {
            using var httpClient = new HttpClient();
            foreach (var header in headers)
            {
                httpClient.DefaultRequestHeaders.Add(header.Key, header.Value);
            }

            var jsonRequestBody = JsonSerializer.Serialize(body);
            var content = new StringContent(jsonRequestBody, System.Text.Encoding.UTF8, "application/json");

            var response = await httpClient.PostAsync(url, content);
            response.EnsureSuccessStatusCode();

            return await response.Content.ReadAsStringAsync();
        }

        private static async Task<string> GetMockFlightDataAsync(string date, string location, string destination)
        {
            var jsonContent = await GetMockDataAsync("flights.json");
            var processedContent = ReplaceDatePlaceholders(jsonContent, DateTime.Now);
            
            // Filter flights by location and/or destination if provided
            if (!string.IsNullOrEmpty(location) || !string.IsNullOrEmpty(destination))
            {
                return FilterFlightsByLocationAndDestination(processedContent, location, destination);
            }
            
            return processedContent;
        }

        private static async Task<string> GetMockHotelDataAsync(string date, string destination)
        {
            var jsonContent = await GetMockDataAsync("hotels.json");
            var processedContent = ReplaceDatePlaceholders(jsonContent, DateTime.Now);
            
            // Filter hotels by destination if provided
            if (!string.IsNullOrEmpty(destination))
            {
                return FilterHotelsByCity(processedContent, destination);
            }
            
            return processedContent;
        }

        private static async Task<string> GetMockDataAsync(string fileName)
        {
            var mockDataPath = Path.Combine("MockData", fileName);
            if (!File.Exists(mockDataPath))
            {
                throw new FileNotFoundException($"Mock data file not found: {mockDataPath}");
            }

            var jsonContent = await File.ReadAllTextAsync(mockDataPath);
            return ReplaceDatePlaceholders(jsonContent, DateTime.Now);
        }

        private static string ReplaceDatePlaceholders(string jsonContent, DateTime baseDateTime)
        {
            // Replace {{NOW_PLUS_DAYS:X}} followed by time (e.g., " 08:30")
            var daysWithTimePattern = @"\{\{NOW_PLUS_DAYS:(\d+)\}\}\s+(\d{2}:\d{2})";
            jsonContent = Regex.Replace(jsonContent, daysWithTimePattern, match =>
            {
                var days = int.Parse(match.Groups[1].Value);
                var timeStr = match.Groups[2].Value;
                var timeParts = timeStr.Split(':');
                var hour = int.Parse(timeParts[0]);
                var minute = int.Parse(timeParts[1]);

                var targetDate = baseDateTime.AddDays(days);
                var targetDateTime = new DateTime(targetDate.Year, targetDate.Month, targetDate.Day, hour, minute, 0);

                return targetDateTime.ToString("yyyy-MM-dd HH:mm");
            });

            // Replace {{NOW_PLUS_HOURS:X}} followed by time (e.g., " 08:15")
            var hoursWithTimePattern = @"\{\{NOW_PLUS_HOURS:(\d+)\}\}\s+(\d{2}:\d{2})";
            jsonContent = Regex.Replace(jsonContent, hoursWithTimePattern, match =>
            {
                var hours = int.Parse(match.Groups[1].Value);
                var timeStr = match.Groups[2].Value;
                var timeParts = timeStr.Split(':');
                var hour = int.Parse(timeParts[0]);
                var minute = int.Parse(timeParts[1]);

                var targetDateTime = baseDateTime.AddHours(hours);
                targetDateTime = new DateTime(targetDateTime.Year, targetDateTime.Month, targetDateTime.Day, hour, minute, 0);

                return targetDateTime.ToString("yyyy-MM-dd HH:mm");
            });

            return jsonContent;
        }

        private static string FilterFlightsByLocationAndDestination(string jsonContent, string location, string destination)
        {
            try
            {
                using var document = JsonDocument.Parse(jsonContent);
                var root = document.RootElement;
                
                if (!root.TryGetProperty("best_flights", out var bestFlightsElement))
                {
                    return jsonContent;
                }

                var filteredFlights = new List<JsonElement>();
                
                foreach (var flight in bestFlightsElement.EnumerateArray())
                {
                    var matchesLocation = string.IsNullOrEmpty(location);
                    var matchesDestination = string.IsNullOrEmpty(destination);

                    if (!string.IsNullOrEmpty(location) && 
                        flight.TryGetProperty("departure_city", out var departureCityElement))
                    {
                        var departureCity = departureCityElement.GetString();
                        if (!string.IsNullOrEmpty(departureCity) && 
                            departureCity.Contains(location, StringComparison.OrdinalIgnoreCase))
                        {
                            matchesLocation = true;
                        }
                    }
                    
                    if (!string.IsNullOrEmpty(destination) && 
                        flight.TryGetProperty("arrival_city", out var arrivalCityElement))
                    {
                        var arrivalCity = arrivalCityElement.GetString();
                        if (!string.IsNullOrEmpty(arrivalCity) && 
                            arrivalCity.Contains(destination, StringComparison.OrdinalIgnoreCase))
                        {
                            matchesDestination = true;
                        }
                    }
                    
                    if (matchesLocation && matchesDestination)
                    {
                        filteredFlights.Add(flight);
                    }
                }

                var filteredResult = new
                {
                    best_flights = filteredFlights.Select(flight => JsonSerializer.Deserialize<object>(flight.GetRawText())).ToArray()
                };

                return JsonSerializer.Serialize(filteredResult, new JsonSerializerOptions 
                { 
                    WriteIndented = false 
                });
            }
            catch
            {
                return jsonContent;
            }
        }

        private static string FilterHotelsByCity(string jsonContent, string city)
        {
            try
            {
                using var document = JsonDocument.Parse(jsonContent);
                var root = document.RootElement;
                
                if (!root.TryGetProperty("hotels", out var hotelsElement))
                {
                    return jsonContent;
                }

                var filteredHotels = new List<JsonElement>();
                
                foreach (var hotel in hotelsElement.EnumerateArray())
                {
                    if (hotel.TryGetProperty("city", out var cityElement))
                    {
                        var hotelCity = cityElement.GetString();
                        
                        if (!string.IsNullOrEmpty(hotelCity) && 
                            hotelCity.Contains(city, StringComparison.OrdinalIgnoreCase))
                        {
                            filteredHotels.Add(hotel);
                        }
                    }
                }

                var filteredResult = new
                {
                    hotels = filteredHotels.Select(hotel => JsonSerializer.Deserialize<object>(hotel.GetRawText())).ToArray()
                };

                return JsonSerializer.Serialize(filteredResult, new JsonSerializerOptions 
                { 
                    WriteIndented = false 
                });
            }
            catch
            {
                return jsonContent;
            }
        }
    }
}
