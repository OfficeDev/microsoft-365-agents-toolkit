#!/bin/bash
set -e

SCRIPT_DIR=$(cd $(dirname "${BASH_SOURCE[0]}") && pwd)
REPO_ROOT_DIR=$(cd $SCRIPT_DIR/../.. && pwd)
TEMPLATE_DIR=$(cd $REPO_ROOT_DIR/templates && pwd)
TEMPLATE_CONFIG_PATH=$REPO_ROOT_DIR/packages/fx-core/src/common/templates-config.json

echo "------ script dir: " $SCRIPT_DIR
echo "------ repo root dir: " $REPO_ROOT_DIR
echo "------ templates dir: " $TEMPLATE_DIR

if [ $1 == 'core-template' ]; then
    echo "sync up templates version in fx-core config"
    node $SCRIPT_DIR/fxcore-sync-up-version.js syncVersion
    git add $TEMPLATE_CONFIG_PATH
    node $SCRIPT_DIR/fxcore-sync-up-version.js updateUseLocalFlag
    git update-index --assume-unchanged $TEMPLATE_CONFIG_PATH
elif [ $1 == 'template-sync' ]; then
    echo "sync up templates deps' version with all the lerna pkgs"
    node $SCRIPT_DIR/sync-version.js
    git add .
fi