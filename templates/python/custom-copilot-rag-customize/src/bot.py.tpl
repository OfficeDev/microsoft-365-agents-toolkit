import os
import sys
import traceback
import json
from dataclasses import asdict
from botbuilder.core import MemoryStorage, TurnContext, CardFactory
from teams import Application, ApplicationOptions, TeamsAdapter
from teams.ai import AIOptions
from teams.ai.models import AzureOpenAIModelOptions, OpenAIModel, OpenAIModelOptions, PromptResponse
from teams.ai.planners import ActionPlanner, ActionPlannerOptions
from teams.ai.prompts import PromptManager, PromptManagerOptions
from teams.state import TurnState, MemoryBase
from teams.feedback_loop_data import FeedbackLoopData
from teams.streaming import StreamingResponse

from my_data_source import MyDataSource

from config import Config

config = Config()

# Create AI components
model: OpenAIModel

{{#useAzureOpenAI}}
model = OpenAIModel(
    AzureOpenAIModelOptions(
        api_key=config.AZURE_OPENAI_API_KEY,
        default_model=config.AZURE_OPENAI_MODEL_DEPLOYMENT_NAME,
        endpoint=config.AZURE_OPENAI_ENDPOINT,
        {{#CEAEnabled}} 
        stream: true,
        {{/CEAEnabled}}
    )
)
{{/useAzureOpenAI}}    
{{#useOpenAI}}
model = OpenAIModel(
    OpenAIModelOptions(
        api_key=config.OPENAI_API_KEY,
        default_model=config.OPENAI_MODEL_NAME,
        {{#CEAEnabled}} 
        stream: true,
        {{/CEAEnabled}}
    )
)
{{/useOpenAI}}

{{#CEAEnabled}}
def end_stream_handler(
    context: TurnContext,
    state: MemoryBase,
    response: PromptResponse[str],
    streamer: StreamingResponse,
):
    if not streamer:
        return

    card = CardFactory.adaptive_card(
        {
            "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
            "version": "1.6",
            "type": "AdaptiveCard",
            "body": [{"type": "TextBlock", "wrap": True, "text": streamer.message}],
        }
    )

    streamer.set_attachments([card])
{{/CEAEnabled}}
    
prompts = PromptManager(PromptManagerOptions(prompts_folder=f"{os.getcwd()}/prompts"))

my_data_source = MyDataSource('local-search')
prompts.add_data_source(my_data_source)

planner = ActionPlanner(
    {{#CEAEnabled}}
    ActionPlannerOptions(model=model, prompts=prompts, default_prompt="chat", start_streaming_message="Loading streaming results...",
                    end_stream_handler=end_stream_handler)
    {{/CEAEnabled}}
    {{^CEAEnabled}}
    ActionPlannerOptions(model=model, prompts=prompts, default_prompt="chat")
    {{/CEAEnabled}}
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

@bot_app.feedback_loop()
async def feedback_loop(_context: TurnContext, _state: TurnState, feedback_loop_data: FeedbackLoopData):
    # Add custom feedback process logic here.
    print(f"Your feedback is:\n{json.dumps(asdict(feedback_loop_data), indent=4)}")