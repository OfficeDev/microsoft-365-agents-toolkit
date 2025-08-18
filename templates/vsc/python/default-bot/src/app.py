"""
Copyright (c) Microsoft Corporation. All rights reserved.
Licensed under the MIT License.
"""

import logging
from os import environ, path
from aiohttp import web
from dotenv import load_dotenv

from microsoft.agents.hosting.aiohttp import CloudAdapter, jwt_authorization_middleware, start_agent_process
from microsoft.agents.authentication.msal import MsalConnectionManager
from microsoft.agents.activity import load_configuration_from_env
from microsoft.agents.hosting.core import Authorization, MemoryStorage, AgentApplication

from bot import teams_bot

# Load environment variables
load_dotenv(path.join(path.dirname(__file__), "./env/.env.dev"))

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
teams_bot.adapter = adapter

# Create routes
routes = web.RouteTableDef()

@routes.post("/api/messages")
async def on_messages(req: web.Request) -> web.Response:
    """Handle incoming messages."""
    try:
        agent: AgentApplication = req.app["agent_app"]
        adapter: CloudAdapter = req.app["adapter"]
        return await start_agent_process(
            req,
            agent,
            adapter,
        )
    except Exception as e:
        logger.error(f"Error processing message: {e}")
        return web.Response(status=500, text=str(e))

# Create the web application
aiohttp_web_app = web.Application(middlewares=jwt_authorization_middleware)
aiohttp_web_app.add_routes(routes)

aiohttp_web_app["agent_configuration"] = connection_manager.get_default_connection_configuration()
aiohttp_web_app["agent_app"] = teams_bot
aiohttp_web_app["adapter"] = adapter

if __name__ == "__main__":
    port = int(environ.get("PORT", 3978))
    client_id = agents_sdk_config.get('client_id', 'Unknown')
    debug = environ.get("DEBUG", "false").lower() == "true"
    
    logger.info(f"Bot Started, listening to port {port} for appId {client_id} debug {debug}")
    
    web.run_app(aiohttp_web_app, host="localhost", port=port)
