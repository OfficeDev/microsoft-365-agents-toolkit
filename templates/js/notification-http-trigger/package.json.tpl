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
        "dev": "func start --javascript --language-worker=\"--inspect=9239\" --port \"3978\" --cors \"*\"",
        "prepare-storage:teamsfx": "azurite --silent --location ./_storage_emulator --debug ./_storage_emulator/debug.log",
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
        "azurite": "^3.16.0",
        "env-cmd": "^10.1.0"
    }
}