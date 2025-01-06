import os
import sys
import traceback

from typing import Any, Dict, List
from botbuilder.core import MemoryStorage, TurnContext, CardFactory, MessageFactory
from teams import Application, ApplicationOptions, TeamsAdapter
from teams.ai import AIOptions
from teams.ai.actions import ActionTurnContext
from teams.ai.models import AzureOpenAIModelOptions, OpenAIModel, OpenAIModelOptions
from teams.ai.planners import ActionPlanner, ActionPlannerOptions
from teams.ai.prompts import PromptManager, PromptManagerOptions
from teams.state import TurnState
from teams.ai.prompts import PromptFunctions, PromptManager, PromptManagerOptions
from teams.ai.tokenizers import Tokenizer
from teams.state import MemoryBase

from config import Config
from state import AppTurnState
from lib.requests_openapi import OpenAPIClient
from lib.adaptive_card_renderer import AdaptiveCardRenderer
import json

config = Config()

# Create AI components
model: OpenAIModel

{{#useAzureOpenAI}}
model = OpenAIModel(
    AzureOpenAIModelOptions(
        api_key=config.AZURE_OPENAI_API_KEY,
        default_model=config.AZURE_OPENAI_MODEL_DEPLOYMENT_NAME,
        endpoint=config.AZURE_OPENAI_ENDPOINT,
    )
)
{{/useAzureOpenAI}}    
{{#useOpenAI}}
model = OpenAIModel(
    OpenAIModelOptions(
        api_key=config.OPENAI_API_KEY,
        default_model=config.OPENAI_MODEL_NAME,
    )
)
{{/useOpenAI}}
    
prompts_folder_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "prompts")
prompts = PromptManager(PromptManagerOptions(prompts_folder=f"{prompts_folder_path}"))

planner = ActionPlanner(
    ActionPlannerOptions(model=model, prompts=prompts, default_prompt="chat")
)

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

@bot_app.error
async def on_error(context: TurnContext, error: Exception):
    # This check writes out errors to console log .vs. app insights.
    # NOTE: In production environment, you should consider logging this to Azure
    #       application insights.
    print(f"\n [on_turn_error] unhandled error: {error}", file=sys.stderr)
    traceback.print_exc()

    # Send a message to the user
    await context.send_activity("The bot encountered an error or bug.")


current_dir = os.path.dirname(os.path.abspath(__file__))
spec_path = os.path.join(current_dir, '../appPackage/apiSpecificationFile/{{OPENAPI_SPEC_PATH}}')
client = OpenAPIClient().load_spec_from_file(spec_path)

@prompts.function("getAction")
async def get_actions(
    _context: TurnContext,
    state: MemoryBase,
    _functions: PromptFunctions,
    _tokenizer: Tokenizer,
    _args: List[str],
):
    action_path = os.path.join(current_dir, 'prompts/chat/actions.json')
    # Read the file content
    with open(action_path, 'r') as action_file:
        action_file_content = action_file.read()

    return action_file_content

# Replace with action code