#!/bin/bash
filePath=packages/vscode-extension/src/releaseBasedFeatureSettings.ts
echo "Update feature settings in $filePath if alpha or preview release"
sed -i -e "s@const shouldEnableTeamsCopilotChatUI = false@const shouldEnableTeamsCopilotChatUI = true@g" $filePath
echo "Prerelease feature setting update done."
