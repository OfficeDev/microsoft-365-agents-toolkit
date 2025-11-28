import { TurnContext } from "botbuilder";
import { AI, TurnState } from "@microsoft/teams-ai";

export async function httpErrorAction(
  context: TurnContext,
  state: TurnState,
  data
): Promise<string> {
  await context.sendActivity("An AI request failed. Please try again later.");
  return AI.StopCommandName;
}

interface WeatherParameters {
  location: string;
  unit?: "c" | "f";
}

export async function getCurrentWeather(
  context: TurnContext,
  state: TurnState,
  parameters: WeatherParameters
): Promise<string> {
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

interface NicknameParameters {
  location: string;
}

export async function getNickname(
  context: TurnContext,
  state: TurnState,
  parameters: NicknameParameters
): Promise<string> {
  const nicknames = {
    "San Francisco": "The Golden City",
    "Los Angeles": "LA",
  };

  return nicknames[parameters.location] ?? `No nickname for ${parameters.location} found`;
}
