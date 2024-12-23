{
    "name": "{{SafeProjectNameLowerCase}}",
    "version": "1.0.0",
    "description": "Microsoft Teams Toolkit Notification Bot Sample",
    "engines": {
        "node": "18 || 20"
    },
    "author": "Microsoft",
    "license": "MIT",
    "scripts": {
        "dev:teamsfx": "env-cmd --silent -f .localConfigs npm run dev",
        "dev:teamsfx:testtool": "env-cmd --silent -f .localConfigs.testTool npm run dev",
        "dev:teamsfx:launch-testtool": "env-cmd --silent -f env/.env.testtool teamsapptester start",
        "dev": "func start --typescript --language-worker=\"--inspect=9239\" --port \"3978\" --cors \"*\"",
        "prepare-storage:teamsfx": "azurite --silent --location ./_storage_emulator --debug ./_storage_emulator/debug.log",
        "watch:teamsfx": "tsc --watch",
        "build": "tsc && shx cp -r ./src/adaptiveCards ./dist/src",
        "watch": "tsc -w",
        "prestart": "npm run build",
        "start": "npx func start",
        "test": "echo \"Error: no test specified\" && exit 1"
    },
    "repository": {
        "type": "git",
        "url": "https://github.com"
    },
    "dependencies": {
        "adaptivecards-templating": "^2.3.1",
        "adaptive-expressions": "^4.23.1",
        "@microsoft/teamsfx": "^3.0.0",
        "botbuilder": "^4.23.1"
    },
    "devDependencies": {
        "@azure/functions": "^3.5.0",
        "@types/json-schema": "^7.0.15",
        "azurite": "^3.16.0",
        "env-cmd": "^10.1.0",
        "ts-node": "^10.4.0",
        "typescript": "^4.4.4",
        "shx": "^0.3.4"
    }
}