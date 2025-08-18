"""
Copyright (c) Microsoft Corporation. All rights reserved.
Licensed under the MIT License.
"""

import os
from dotenv import load_dotenv

load_dotenv(os.path.join(os.path.dirname(__file__), "./env/.env.dev"))

class Config:
    """Bot Configuration"""
    
    PORT = int(os.environ.get("PORT", 3978))
    BOT_ID = os.environ.get("BOT_ID", "")
    BOT_PASSWORD = os.environ.get("BOT_PASSWORD", "")
    BOT_TYPE = os.environ.get("BOT_TYPE", "")
    BOT_TENANT_ID = os.environ.get("BOT_TENANT_ID", "")
    
    # Debug mode
    DEBUG = os.environ.get("DEBUG", "false").lower() == "true"
