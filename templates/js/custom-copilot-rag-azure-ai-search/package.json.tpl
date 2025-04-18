{
    "name": "{{SafeProjectNameLowerCase}}",
    "version": "1.0.0",
    "msteams": {
        "teamsAppId": null
    },
    "description": "Microsoft 365 Agents Toolkit RAG Bot Sample with Azure AI Search and Teams AI Library",
    "engines": {
        "node": "18 || 20 || 22"
    },
    "author": "Microsoft",
    "license": "MIT",
    "main": "./src/index.js",
    "scripts": {
        "dev:teamsfx": "env-cmd --silent -f .localConfigs npm run dev",
        "dev:teamsfx:testtool": "env-cmd --silent -f .localConfigs.testTool npm run dev",
        "dev:teamsfx:launch-testtool": "env-cmd --silent -f env/.env.testtool teamsapptester start",
        "dev": "nodemon --inspect=9239 --signal SIGINT ./src/index.js",
        "start": "node ./src/index.js",
        "test": "echo \"Error: no test specified\" && exit 1",
        "watch": "nodemon --exec \"npm run start\"",
        "indexer:create": "env-cmd --silent -f env/.env.testtool.user node ./src/indexers/setup.js",
        "indexer:delete": "env-cmd --silent -f env/.env.testtool.user node ./src/indexers/delete.js"
    },
    "repository": {
        "type": "git",
        "url": "https://github.com"
    },
    "dependencies": {
        "@azure/search-documents": "^12.0.0",
        "@microsoft/teams-ai": "^1.5.3",
        "botbuilder": "^4.23.1",
        "express": "^5.0.1"
    },
    "devDependencies": {
        "env-cmd": "^10.1.0",
        "nodemon": "^3.1.7"
    }
}