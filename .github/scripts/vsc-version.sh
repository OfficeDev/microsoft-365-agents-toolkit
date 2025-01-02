#!/bin/bash
SCRIPT_DIR=$(cd $(dirname "${BASH_SOURCE[0]}") && pwd)
REPO_ROOT_DIR=$(cd $SCRIPT_DIR/../.. && pwd)
VSC_DIR=$(cd $REPO_ROOT_DIR/packages/vscode-extension && pwd)
VSC_PACKAGE_JSON_DIR=$VSC_DIR/package.json
VERSION=$(jq -r .version $VSC_PACKAGE_JSON_DIR)
echo '-----------------' $VERSION
# get minor version
MINOR_VER=$(echo $VERSION | awk -F. '{print $2}')
DATE_WITH_TIME=`date "+%Y%m%d%H"`
# prerelease version should set minor version to odd number, and set patch version as timestamp.
if [ "$PREID" == "preview" ]; then
    if [ $((MINOR_VER%2)) -eq 0 ]; then
        echo "Need to bump up version with even minor version for preview"
        VERSION=$(echo ${VERSION%-*} | awk -v val=$DATE_WITH_TIME -F. '/[0-9]+\./{$2++;$3=val;print}' OFS=.)
    else
        echo "Need to set patch version as timestamp for preview"
        VERSION=$(echo ${VERSION%-*} | awk -v val=$DATE_WITH_TIME -F. '/[0-9]+\./{$3=val;print}' OFS=.)
    fi
    echo '=====================' $VERSION
    # update the package.json file
    jq --arg VERSION "$VERSION" '.version=$VERSION' package.json > tmp.$$.json
    mv tmp.$$.json $VSC_PACKAGE_JSON_DIR
fi