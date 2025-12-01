const { AI } = require("@microsoft/teams-ai");

async function httpErrorAction(context, state, data) {
  await context.sendActivity("An AI request failed. Please try again later.");
  return AI.StopCommandName;
}

async function getCurrentWeather(context, state, parameters) {
  const weatherData = {
    "San Francisco, CA": {
      f: "71.6F",
      c: "22C",
    },
    "Los Angeles": {
      f: "75.2F",
      c: "24C",
    },
  };

  if (weatherData[parameters.location] === undefined) {
    return `No weather data for ${parameters.location} found`;
  }

  return weatherData[parameters.location][parameters.unit ?? "f"];
}

async function getNickname(context, state, parameters) {
  const nicknames = {
    "San Francisco": "The Golden City",
    "Los Angeles": "LA",
  };

  return nicknames[parameters.location] ?? `No nickname for ${parameters.location} found`;
}

module.exports = {
  httpErrorAction,
  getCurrentWeather,
  getNickname,
};
