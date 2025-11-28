using Microsoft.Recognizers.Text;
using Microsoft.Recognizers.Text.DateTime;

namespace {{SafeProjectName}}.Utils
{
    public static class TimeRangeUtils
    {
        public static (DateTime From, DateTime To)? ExtractTimeRange(
            string timePhrase,
            ILogger logger
        )
        {
            if (string.IsNullOrWhiteSpace(timePhrase))
            {
                return null;
            }

            try
            {
                var results = DateTimeRecognizer.RecognizeDateTime(timePhrase, Culture.English);
                if (results == null || results.Count == 0)
                {
                    logger.LogWarning(
                        "Could not parse time phrase: \"{Phrase}\"",
                        timePhrase
                    );
                    return null;
                }

                var now = DateTime.UtcNow;
                var resolution = results[0].Resolution;

                if (resolution == null || !resolution.ContainsKey("values"))
                {
                    logger.LogWarning(
                        "No resolution values found for: \"{Phrase}\"",
                        timePhrase
                    );
                    return null;
                }

                if (resolution["values"] is not List<Dictionary<string, string>> values || values.Count == 0)
                {
                    logger.LogWarning(
                        "Empty resolution values for: \"{Phrase}\"",
                        timePhrase
                    );
                    return null;
                }

                var firstValue = values[0];
                if (!firstValue.ContainsKey("type"))
                {
                    return null;
                }

                var type = firstValue["type"];
                if (type == "daterange" || type == "datetimerange")
                {
                    if (
                        firstValue.TryGetValue("start", out var startValue)
                        && firstValue.TryGetValue("end", out var endValue)
                        && DateTime.TryParse(startValue, out var start)
                        && DateTime.TryParse(endValue, out var end)
                    )
                    {
                        return (start, end);
                    }
                }
                else if (type == "date")
                {
                    if (
                        firstValue.TryGetValue("value", out var dateValue)
                        && DateTime.TryParse(dateValue, out var date)
                    )
                    {
                        return (date.Date, date.Date.AddDays(1).AddSeconds(-1));
                    }
                }
                else if (type == "datetime")
                {
                    if (
                        firstValue.TryGetValue("value", out var datetimeValue)
                        && DateTime.TryParse(datetimeValue, out var dateTime)
                    )
                    {
                        return dateTime <= now ? (dateTime, now) : (now, dateTime);
                    }
                }
            }
            catch (Exception ex)
            {
                logger.LogError(ex, "Error parsing time phrase: {Phrase}", timePhrase);
            }

            logger.LogWarning(
                "Could not extract time range from resolution for: \"{Phrase}\"",
                timePhrase
            );
            return null;
        }
    }
}
