"""
Copyright (c) Microsoft Corporation. All rights reserved.
Licensed under the MIT License.
"""

import logging
from os import environ
from aiohttp import web
from dotenv import load_dotenv

from microsoft.agents.hosting.aiohttp import CloudAdapter
from microsoft.agents.authentication.msal import MsalConnectionManager
from microsoft.agents.activity import load_configuration_from_env
from microsoft.agents.hosting.core import Authorization, MemoryStorage

from bot import teams_bot

# Load environment variables
load_dotenv()

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

# Create authorization
authorization = Authorization(storage, connection_manager, **agents_sdk_config)

# Create routes
routes = web.RouteTableDef()

@routes.post("/api/messages")
async def on_messages(req: web.Request) -> web.Response:
    """Handle incoming messages."""
    try:
        # Process the request through the adapter
        return await adapter.process(req, teams_bot.run)
    except Exception as e:
        logger.error(f"Error processing message: {e}")
        return web.Response(status=500, text=str(e))

# Create the web application
app = web.Application()
app.add_routes(routes)

if __name__ == "__main__":
    port = int(environ.get("PORT", 3978))
    client_id = agents_sdk_config.get('client_id', 'Unknown')
    debug = environ.get("DEBUG", "false").lower() == "true"
    
    logger.info(f"Bot Started, listening to port {port} for appId {client_id} debug {debug}")
    
    web.run_app(app, host="localhost", port=port)
