#!/bin/bash
LATEST_TAG=$(curl -s https://api.github.com/repos/OfficeDev/microsoft-365-agents-toolkit/releases | grep -o 
'
"tag_name": "microsoft-365-agents-playground@[^"]*"
'
 | head -1 | cut -d
'
 -f4) && curl -L -o agentsplayground.zip "https://github.com/OfficeDev/microsoft-365-agents-toolkit/releases/download/${LATEST_TAG}/agentsplayground-linux-x64.zip" && python3 -m zipfile -e agentsplayground.zip . && chmod +x agentsplayground
