"""
Copyright (c) Microsoft Corporation. All rights reserved.
Licensed under the MIT License.
"""

import logging
import json

from os import environ, path
from dotenv import load_dotenv

from microsoft_agents.authentication.msal import MsalConnectionManager
from microsoft_agents.hosting.core import (
    AgentApplication, 
    TurnState, 
    TurnContext, 
    MemoryStorage
)
from microsoft_agents.activity import ActivityTypes, load_configuration_from_env
from microsoft_agents.hosting.aiohttp import CloudAdapter

# Load environment variables
load_dotenv(path.join(path.dirname(__file__), "../env/.env.dev"))

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Load configuration from environment
agents_sdk_config = load_configuration_from_env(environ)

# Create storage and connection manager
storage = MemoryStorage()
connection_manager = MsalConnectionManager(**agents_sdk_config)

# Create adapter
adapter = CloudAdapter(connection_manager=connection_manager)

# Create the agent application
teams_bot = AgentApplication[TurnState](
    storage=storage
)

@teams_bot.message("/reset")
async def on_reset_message(context: TurnContext, state: TurnState):
    """Reset conversation state when user types /reset."""
    state.clear()
    await context.send_activity("Ok I've deleted the current conversation state.")

@teams_bot.message("/count")
async def on_count_message(context: TurnContext, state: TurnState):
    """Show current message count when user types /count."""
    count = state.conversation.get_value("count", 0, target_cls=int)
    await context.send_activity(f"The count is {count}")

@teams_bot.message("/diag")
async def on_diag_message(context: TurnContext, state: TurnState):
    """Show diagnostic information when user types /diag."""
    await state.load(context, storage)
    activity_json = json.dumps(context.activity.model_dump(by_alias=True, exclude_unset=True), indent=2, default=str)
    await context.send_activity(f"```json\n{activity_json}\n```")

@teams_bot.message("/state")
async def on_state_message(context: TurnContext, state: TurnState):
    """Show current state when user types /state."""
    await state.load(context, storage)
    state_dict = {
        "conversation": state.conversation.__dict__ if state.conversation else None,
        "user": state.user.__dict__ if state.user else None,
        "temp": state.temp.__dict__ if state.temp else None
    }
    state_json = json.dumps(state_dict, indent=2, default=str)
    await context.send_activity(f"```json\n{state_json}\n```")

@teams_bot.message("/runtime")
async def on_runtime_message(context: TurnContext, state: TurnState):
    """Show runtime information when user types /runtime."""
    import sys
    try:
        from microsoft_agents.hosting.core import __version__ as agents_version
    except ImportError:
        agents_version = "unknown"
    
    runtime_info = {
        "python_version": sys.version,
        "agents_sdk_version": agents_version
    }
    runtime_json = json.dumps(runtime_info, indent=2)
    await context.send_activity(f"```json\n{runtime_json}\n```")

@teams_bot.conversation_update("membersAdded")
async def on_members_added(context: TurnContext, state: TurnState):
    """Welcome new members when they're added to the conversation."""
    try:
        from microsoft_agents.hosting.core import __version__ as agents_version
    except ImportError:
        agents_version = "unknown"
    
    await context.send_activity(
        f"Hi there! I'm an echo bot running on Agents SDK version {agents_version} that will echo what you said to me."
    )

@teams_bot.activity(ActivityTypes.message)
async def on_message_activity(context: TurnContext, state: TurnState):
    """Handle all message activities - echo back what the user said with a count."""
    
    # Increment count
    count = state.conversation.get_value("count", lambda: 0, target_cls=int)
    state.conversation.set_value("count", count + 1)
    
    await state.save(context, storage)

    # Echo back user's message with count
    await context.send_activity(f"[{state.conversation.get_value('count', target_cls=int)}] you said: {context.activity.text}")

# Additional pattern matching examples (similar to TypeScript version)
# Note: These may need to be adjusted based on the actual Python API
async def custom_message_filter(context: TurnContext) -> bool:
    """Custom filter function to match message activities."""
    return context.activity.type == "message"

@teams_bot.activity(custom_message_filter)
async def on_custom_filter_message(context: TurnContext, state: TurnState):
    """Handle activities that match a custom filter function."""
    await context.send_activity(f"Matched function: {context.activity.type}")
