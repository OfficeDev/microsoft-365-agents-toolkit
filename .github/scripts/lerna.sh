#!/bin/bash
VAR=$@
echo ============= Inputs: $VAR ================
stringarray=($VAR)
git update-index --assume-unchanged "pnpm-workspace.yaml"
if [ ${#stringarray[@]} -eq 0 ]; then
    echo "for all the pkgs except mcp-server"
    # Remove the packages/mcp-server line from the workspace
    sed -i '/^[[:space:]]*-[[:space:]]*"packages\/mcp-server"/d' pnpm-workspace.yaml
    cat pnpm-workspace.yaml
    exit 0
else
    content=$(jq ".common" .github/scripts/lernaDeps.json)
    for i in "${stringarray[@]}"
    do :
        echo package name: $i
        if [ $(jq --arg v "$i" 'has($v)' .github/scripts/lernaDeps.json) == 'false' ]; then
            echo "Get Error Inputs:" $i
            exit -1
        fi
        pkgContent=$(jq --arg a "$i" '.[$a]' -r .github/scripts/lernaDeps.json)
        content=$(jq --argjson arr1 "$content" --argjson arr2 "$pkgContent" -n '$arr1 + $arr2 | unique')
    done
    content=$(echo $content|jq -r '.[]')
    echo ======== deps: $content ==========
    echo "packages:" > tmp.$$.yaml
    for key in $content
    do
        echo "- $key" >> tmp.$$.yaml
    done
    mv tmp.$$.yaml pnpm-workspace.yaml
    cat pnpm-workspace.yaml
fi