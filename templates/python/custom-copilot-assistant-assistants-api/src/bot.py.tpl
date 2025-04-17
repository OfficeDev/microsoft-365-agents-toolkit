import os
import sys
import traceback
import json
from typing import Any, Dict, Optional
from dataclasses import asdict

from botbuilder.core import MemoryStorage, TurnContext
from teams import Application, ApplicationOptions, TeamsAdapter
from teams.ai import AIOptions
from teams.ai.planners import AssistantsPlanner, OpenAIAssistantsOptions, AzureOpenAIAssistantsOptions
from teams.state import TurnState
from teams.feedback_loop_data import FeedbackLoopData

from config import Config

config = Config()

{{#useOpenAI}}
planner = AssistantsPlanner[TurnState](
    OpenAIAssistantsOptions(api_key=config.OPENAI_API_KEY, assistant_id=config.OPENAI_ASSISTANT_ID)
)
{{/useOpenAI}}
{{#useAzureOpenAI}}
planner = AssistantsPlanner[TurnState](
    AzureOpenAIAssistantsOptions(
        api_key=config.AZURE_OPENAI_API_KEY,
        endpoint=config.AZURE_OPENAI_ENDPOINT,
        default_model=config.AZURE_OPENAI_MODEL_DEPLOYMENT_NAME,
        assistant_id=config.AZURE_OPENAI_ASSISTANT_ID)
)
{{/useAzureOpenAI}}

# Define storage and application
storage = MemoryStorage()
bot_app = Application[TurnState](
    ApplicationOptions(
        bot_app_id=config.APP_ID,
        storage=storage,
        adapter=TeamsAdapter(config),
        ai=AIOptions(planner=planner, enable_feedback_loop=True),
    )
)
    
@bot_app.ai.action("getCurrentWeather")
async def get_current_weather(context: TurnContext, state: TurnState):
    weatherData = {
        'San Francisco, CA': {
            'f': '71.6F',
            'c': '22C',
        },
        'Los Angeles': {
            'f': '75.2F',
            'c': '24C',
        },
    }
    location = context.data.get("location")
    if not weatherData.get(location):
        return f"No weather data for ${location} found"
    
    return weatherData[location][context.data.get("unit") if context.data.get("unit") else 'f']

@bot_app.ai.action("getNickname")
async def get_nickname(context: TurnContext, state: TurnState):
    nicknames = {
        'San Francisco, CA': 'The Golden City',
        'Los Angeles': 'LA',
    }
    location = context.data.get("location")
    
    return nicknames.get(location) if nicknames.get(location) else f"No nickname for ${location} found"

@bot_app.error
async def on_error(context: TurnContext, error: Exception):
    # This check writes out errors to console log .vs. app insights.
    # NOTE: In production environment, you should consider logging this to Azure
    #       application insights.
    print(f"\n [on_turn_error] unhandled error: {error}", file=sys.stderr)
    traceback.print_exc()

    # Send a message to the user
    await context.send_activity("The agent encountered an error or bug.")

@bot_app.feedback_loop()
async def feedback_loop(_context: TurnContext, _state: TurnState, feedback_loop_data: FeedbackLoopData):
    # Add custom feedback process logic here.
    print(f"Your feedback is:\n{json.dumps(asdict(feedback_loop_data), indent=4)}")