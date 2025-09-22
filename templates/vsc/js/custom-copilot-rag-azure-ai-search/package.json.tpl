{
    "name": "{{SafeProjectNameLowerCase}}",
    "version": "1.0.0",
    "msteams": {
        "teamsAppId": null
    },
    "description": "Microsoft 365 Agents Toolkit RAG Bot Sample with Azure AI Search and Teams AI Library",
    "engines": {
        "node": "20 || 22"
    },
    "author": "Microsoft",
    "license": "MIT",
    "main": "./src/index.js",
    "scripts": {
        "dev:teamsfx": "env-cmd --silent -f .localConfigs npm run dev",
        "dev:teamsfx:testtool": "env-cmd --silent -f .localConfigs.playground npm run dev",
        "dev:teamsfx:launch-testtool": "env-cmd --silent -f env/.env.playground teamsapptester start",
        "dev": "nodemon --inspect=9239 --signal SIGINT ./src/index.js",
        "start": "node ./src/index.js",
        "test": "echo \"Error: no test specified\" && exit 1",
        "watch": "nodemon --exec \"npm run start\"",
        "indexer:create": "env-cmd --silent -f env/.env.playground.user node ./src/indexers/setup.js",
        "indexer:delete": "env-cmd --silent -f env/.env.playground.user node ./src/indexers/delete.js"
    },
    "repository": {
        "type": "git",
        "url": "https://github.com"
    },
    "dependencies": {
        "@azure/identity": "^4.11.1",
        "@azure/search-documents": "^12.0.0",
        "@microsoft/teams.apps": "^2.0.0",
        "@microsoft/teams.ai": "^2.0.0",
        "@microsoft/teams.openai": "^2.0.0",
        "@microsoft/teams.common": "^2.0.0"
    },
    "devDependencies": {
        "env-cmd": "^10.1.0",
        "nodemon": "^3.1.7"
    }
}